# Loading external assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$AutoStart = 1     #Set to 1 to run on launch

#### Definitions ####
$Version = "2.0.0"
$AppDataFolder = "$env:APPDATA\OnIT\AssignedAlerts"
$ConfigFile = "AssignedAlerts.config"
$FullConfigPath = "$AppDataFolder\$ConfigFile"
$timer = New-Object System.Windows.Forms.Timer
$URL = "https://app.atera.com/api/v3/tickets?page=1&itemsInPage=50&ticketStatus=Send%20To%20Tech"

$Form1 = New-Object System.Windows.Forms.Form
$button_start = New-Object System.Windows.Forms.Button
$button_pause = New-Object System.Windows.Forms.Button
$button_save = New-Object System.Windows.Forms.Button
$label1 = New-Object System.Windows.Forms.Label
$label2 = New-Object System.Windows.Forms.Label
$label3 = New-Object System.Windows.Forms.Label
$numericUpDown1 = New-Object System.Windows.Forms.NumericUpDown
$label4 = New-Object System.Windows.Forms.Label
$table_tech = New-Object System.Windows.Forms.DataGridView
$label5 = New-Object System.Windows.Forms.Label
$label6 = New-Object System.Windows.Forms.Label
$label7 = New-Object System.Windows.Forms.Label
$tb_ateraAPI = New-Object System.Windows.Forms.TextBox
$tb_status1 = New-Object System.Windows.Forms.TextBox
$tb_status2 = New-Object System.Windows.Forms.TextBox
$tb_twilioSID = New-Object System.Windows.Forms.TextBox
$tb_twilioToken = New-Object System.Windows.Forms.TextBox
$tb_twilioNum = New-Object System.Windows.Forms.TextBox
$label8 = New-Object System.Windows.Forms.Label
$tb_console = New-Object System.Windows.Forms.TextBox
$label9 = New-Object System.Windows.Forms.Label



Function FirstRun {
    if (Test-Path -Path $AppDataFolder) {
    }
    else {
        New-Item -Path "$env:APPDATA\OnIT" -Name "AssignedAlerts" -ItemType "directory" | out-null
    }

    if (Test-Path -Path "$FullConfigPath") {
    }
    else {
        New-Item -Path "$AppDataFolder" -Name "$ConfigFile" -ItemType File | out-null
        "AteraAPIKey:Enter Key;Status1:Status1;Status2:Status2;Pause:20;TwilioSID:Enter SID;TwilioToken:Enter Token;TwilioNumber:Enter Number;Tech:Example1,+12223334444;Tech:Example2,+19998887777;" | add-content -path "$FullConfigPath"
    }
}



function SaveConfig {
    $AteraTokenNew = $tb_ateraAPI.Text
    $TwilioTokenNew = $tb_twilioToken.Text
    $TwilioSIDNew = $tb_twilioSID.Text
    $TwilioNumberNew = $tb_twilioNum.Text
    $RefreshNew = $numericUpDown1.Text
    $Status1New = $tb_status1.Text
    $Status2New = $tb_status2.Text
   
    #Save Techs and Numbers
    $RowCount = $table_tech.RowCount
    $RowCount -= 2
    $Main = 1
    $Techs = @()
    While ($Main -eq 1){
        $Name = $table_tech[0,$RowCount].Value
        $Number = $table_tech[1,$RowCount].Value
        $Techs += "Tech:$Name,$Number"
        $RowCount -= 1
        If ($RowCount -eq -1){
            $Main = 0
        }
    }
    $Techs = $Techs | Sort-Object
    $AllTechs = ""
    Foreach ($TechNum in $Techs){
        $AllTechs += "$TechNum;"
    }

    $configNew = "AteraAPIKey:$AteraTokenNew;Status1:$Status1New;Status2:$Status2New;Pause:$RefreshNew;TwilioSID:$TwilioSIDNew;TwilioToken:$TwilioTokenNew;TwilioNumber:$TwilioNumberNew;$AllTechs"

    Remove-Item -Path $FullConfigPath -Force
    New-Item -Path "$AppDataFolder" -Name "$ConfigFile" -ItemType File | out-null
    "$configNew" | add-content -path "$FullConfigPath"
}


Function SendText{
    param(
        [String]$techPhone,
        [String]$textBody
    )
    ### Send Text Message ###
    $sid = $tb_twilioSID.Text
    $token = $tb_twilioToken.Text
    $number = $tb_twilioNum.Text
    

    # Twilio API endpoint and POST params
    $url = "https://api.twilio.com/2010-04-01/Accounts/$sid/Messages.json"
    $params = @{ To = $techPhone; From = $number; Body = $textBody }

    # Create a credential object for HTTP basic auth
    $p = $token | ConvertTo-SecureString -asPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($sid, $p)

    # Make API request, selecting JSON properties from response
    Invoke-WebRequest $url -Method Post -Credential $credential -Body $params -UseBasicParsing |
    ConvertFrom-Json | Select sid, body
}


Function LoadConfig {
    $Config = Get-Content -Path "$FullConfigPath"

    $ConfigSplit = $config.Split(";")
    $AteraAPIKey = $ConfigSplit[0].Split(":")[1]
    $Status1 = $ConfigSplit[1].Split(":")[1]
    $Status2 = $ConfigSplit[2].Split(":")[1]
    $PauseTime = $ConfigSplit[3].Split(":")[1]
    $TwilioSID = $ConfigSplit[4].Split(":")[1]
    $TwilioToken = $ConfigSplit[5].Split(":")[1]
    $TwilioNumber = $ConfigSplit[6].Split(":")[1]

    $tb_ateraAPI.Text = $AteraAPIKey
    $tb_twilioToken.Text = $TwilioToken
    $tb_twilioSID.Text = $TwilioSID
    $tb_twilioNum.Text = $TwilioNumber
    $tb_status1.Text = $Status1
    $tb_status2.Text = $Status2




    #Update Interval Timers
    $numericUpDown1.Value = $PauseTime
    $timer1MS = [int]$PauseTime*1000
    $timer.Interval = $timer1MS

    


    #Load Tech List
    $table_tech.Rows.Clear();
    $Tech_List = @()
    foreach ($tech in $ConfigSplit){
        If ($tech -like "Tech:*"){
            $techName = $tech.Split(",")[0]
            $techName = $techName.Split(":")[1]
            $techNumber = $tech.Split(",")[1]
            $table_tech.Rows.Add("$techName","$techNumber")
            $Tech_List += $techName
        }
        $table_tech.Refresh();
    }
}


function consoleUpdate {
    param(
        [String]$text
    )
    $consoleTime = Get-Date -f "HH:mm:ss"
    $tb_console.AppendText("$consoleTime - $text`r`n")


}



function GetPhoneNumber{
    param(
        [String]$tech
    )

    $RowCount = $table_tech.RowCount
    $RowCount -= 2
    $Main = 1
    $Techs = @{}
    While ($Main -eq 1){
        $Name = $table_tech[0,$RowCount].Value
        $Number = $table_tech[1,$RowCount].Value
        $Techs += @{"$Name" = "$Number"}
        $RowCount -= 1
        If ($RowCount -eq -1){
            $Main = 0
        }
    }

    $PhoneNumber = $Techs.Get_Item("$tech")
    return $PhoneNumber

}


function startTimer { 
   $timer.start()
}

function stopTimer {
    $timer.Stop()
    $timer.Dispose()
}


Function AutoRun {
    $button_pause.Enabled = $true
    $button_start.Enabled = $false
    startTimer
}


######################## Main ###################

$timer.add_tick({
    #Start Main Program
    $AteraAPIKey = $tb_ateraAPI.Text
    $Status = $tb_status1.Text
    $Status1 = $Status.Replace(" ","%20")
    $URL = "https://app.atera.com/api/v3/tickets?page=1&itemsInPage=50&ticketStatus=$Status1"


    #Checking Atera for tickets
    consoleUpdate "Checking Atera for tickets"
    $Tickets = ""
    $Response = ""

    $Response = Invoke-RestMethod -Method Get -Uri $URL -Header @{ "X-Api-Key" = $ateraApiKey }
    $Tickets = $Response.items

    If (! $Tickets){
        consoleUpdate "No status set to $Status"
    }
    Else{
        #Get tech assigned for each ticket
        Foreach ($Ticket in $Tickets){

            $tech = ""
            $TicketID = ""
            $techphone = ""
            $textBody = ""
            $tech = $Ticket.TechnicianFullName
            $TicketID = $Ticket.TicketID



            #Get techs phone number
            $techphone = GetPhoneNumber -tech $tech



            #Send Text
            consoleUpdate "Sending Text to $tech for ticket number $TicketID"
            $textBody = "You have been assigned ticket # $TicketID"
            SendText -techPhone $techphone -textBody $textBody


            #Set status to Assigned
            $body = @{
                "TicketStatus"="Assigned"
            } | ConvertTo-Json

            $header = @{
                "Content-Type"="application/json"
                "Accept"="application/json"
                "X-API-KEY"="$ateraApiKey"
            }

            consoleUpdate "Setting $TicketID status to assigned"
            Invoke-RestMethod -Uri "https://app.atera.com/api/v3/tickets/$TicketID" -Method 'Post' -Body $body -Headers $header

        }
    }


})


#### Buttons ####
# button_start
$button_start.Location = New-Object System.Drawing.Point(320, 494)
$button_start.Name = "button_start"
$button_start.Size = New-Object System.Drawing.Size(75, 30)
$button_start.TabIndex = 20
$button_start.Text = "Start"
$button_start.UseVisualStyleBackColor = $true
# I added this:
$button_start.add_click({
    $button_pause.Enabled  = $true
    $button_start.Enabled  = $false
    LoadConfig
    startTimer
    consoleUpdate "Starting"

})

# button_pause
$button_pause.Location = New-Object System.Drawing.Point(225, 494)
$button_pause.Name = "button_pause"
$button_pause.Size = New-Object System.Drawing.Size(75, 30)
$button_pause.TabIndex = 19
$button_pause.Text = "Pause"
$button_pause.UseVisualStyleBackColor = $true
$button_pause.Enabled = $false
# I added this:
$button_pause.add_click({
    $button_start.Enabled   = $true
    $button_pause.Enabled   = $false
    stopTimer
    consoleUpdate "Pausing"
})

# button_save
$button_save.Location = New-Object System.Drawing.Point(122, 494)
$button_save.Name = "button_save"
$button_save.Size = New-Object System.Drawing.Size(75, 30)
$button_save.TabIndex = 18
$button_save.Text = "Save"
$button_save.UseVisualStyleBackColor = $true
$button_save.add_click({
    $tb_console.AppendText("Saving`r`n")
    SaveConfig
})


#### Labels ####

# label1
$label1.AutoSize = $true
$label1.Location = New-Object System.Drawing.Point(118, 79)
$label1.Name = "label1"
$label1.Size = New-Object System.Drawing.Size(108, 20)
$label1.TabIndex = 2
$label1.Text = "Atera API Key"

# label2
$label2.AutoSize = $true
$label2.Location = New-Object System.Drawing.Point(96, 118)
$label2.Name = "label2"
$label2.Size = New-Object System.Drawing.Size(130, 20)
$label2.TabIndex = 4
$label2.Text = "Status to look for"

# label3
$label3.AutoSize = $true
$label3.Location = New-Object System.Drawing.Point(80, 154)
$label3.Name = "label3"
$label3.Size = New-Object System.Drawing.Size(146, 20)
$label3.TabIndex = 6
$label3.Text = "Status to update to"

# label4
$label4.AutoSize = $true
$label4.Location = New-Object System.Drawing.Point(53, 200)
$label4.Name = "label4"
$label4.Size = New-Object System.Drawing.Size(173, 20)
$label4.TabIndex = 8
$label4.Text = "Pause between checks"

# label5
$label5.AutoSize = $true
$label5.Location = New-Object System.Drawing.Point(147, 243)
$label5.Name = "label5"
$label5.Size = New-Object System.Drawing.Size(79, 20)
$label5.TabIndex = 11
$label5.Text = "Twilio SID"

# label6
$label6.AutoSize = $true
$label6.Location = New-Object System.Drawing.Point(131, 288)
$label6.Name = "label6"
$label6.Size = New-Object System.Drawing.Size(95, 20)
$label6.TabIndex = 13
$label6.Text = "Twilio Token"

# label7
$label7.AutoSize = $true
$label7.Location = New-Object System.Drawing.Point(119, 328)
$label7.Name = "label7"
$label7.Size = New-Object System.Drawing.Size(107, 20)
$label7.TabIndex = 15
$label7.Text = "Twilio Number"

# label8
$label8.AutoSize = $true
$label8.Location = New-Object System.Drawing.Point(316, 200)
$label8.Name = "label8"
$label8.Size = New-Object System.Drawing.Size(79, 20)
$label8.TabIndex = 10
$label8.Text = "(seconds)"

# label9
$label9.AutoSize = $true
$label9.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 16,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point, 0)
$label9.Location = New-Object System.Drawing.Point(170, 24)
$label9.Name = "label9"
$label9.Size = New-Object System.Drawing.Size(140, 37)
$label9.TabIndex = 1
$label9.Text = "Settings"


#### Counter ####

# numericUpDown1
$numericUpDown1.Location = New-Object System.Drawing.Point(232, 198)
$numericUpDown1.Name = "numericUpDown1"
$numericUpDown1.Size = New-Object System.Drawing.Size(78, 26)
$numericUpDown1.TabIndex = 9
$numericUpDown1.Maximum = 300
$numericUpDown1.Minimum = 1



#### Tech Table ####

# table_tech
$table_tech.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$table_tech.Location = New-Object System.Drawing.Point(576, 12)
$table_tech.Name = "table_tech"
$table_tech.Size = New-Object System.Drawing.Size(414, 329)
$table_tech.TabIndex = 21
# I added this:
$table_tech.RowHeadersVisible = $false
$table_tech.AutoSizeColumnsMode = 'Fill'
$table_tech.AllowUserToResizeRows = $true
$table_tech.selectionmode = 'FullRowSelect'
$table_tech.MultiSelect = $false
$table_tech.AllowUserToAddRows = $true
$table_tech.ReadOnly = $false
$table_tech.ColumnCount = 2
$table_tech.ColumnHeadersVisible = $true
$table_tech.Columns[0].Name = "Tech"
$table_tech.Columns[1].Name = "Number"
$table_tech.Sort($table_tech.Columns['Tech'],'Ascending')


#### Text Boxes ####

# tb_ateraAPI
$tb_ateraAPI.Location = New-Object System.Drawing.Point(232, 76)
$tb_ateraAPI.Name = "tb_ateraAPI"
$tb_ateraAPI.Size = New-Object System.Drawing.Size(181, 26)
$tb_ateraAPI.TabIndex = 3

# tb_status1
$tb_status1.Location = New-Object System.Drawing.Point(232, 115)
$tb_status1.Name = "tb_status1"
$tb_status1.Size = New-Object System.Drawing.Size(181, 26)
$tb_status1.TabIndex = 5

# tb_status2
$tb_status2.Location = New-Object System.Drawing.Point(232, 151)
$tb_status2.Name = "tb_status2"
$tb_status2.Size = New-Object System.Drawing.Size(181, 26)
$tb_status2.TabIndex = 7

# tb_twilioSID
$tb_twilioSID.Location = New-Object System.Drawing.Point(232, 240)
$tb_twilioSID.Name = "tb_twilioSID"
$tb_twilioSID.Size = New-Object System.Drawing.Size(181, 26)
$tb_twilioSID.TabIndex = 12

# tb_twilioToken
$tb_twilioToken.Location = New-Object System.Drawing.Point(232, 282)
$tb_twilioToken.Name = "tb_twilioToken"
$tb_twilioToken.Size = New-Object System.Drawing.Size(181, 26)
$tb_twilioToken.TabIndex = 14

# tb_twilioNum
$tb_twilioNum.Location = New-Object System.Drawing.Point(232, 325)
$tb_twilioNum.Name = "tb_twilioNum"
$tb_twilioNum.Size = New-Object System.Drawing.Size(181, 26)
$tb_twilioNum.TabIndex = 16
$tb_twilioNum.Text = "+1xxxxxxxxxx"

# tb_console
$tb_console.Location = New-Object System.Drawing.Point(576, 347)
$tb_console.Multiline = $true
$tb_console.Name = "tb_console"
$tb_console.ReadOnly = $true
$tb_console.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$tb_console.Size = New-Object System.Drawing.Size(414, 200)
$tb_console.TabIndex = 22


#### Main Form ####

# Form1
$Form1.ClientSize = New-Object System.Drawing.Size(1002, 559)
$Form1.Controls.Add($label9)
$Form1.Controls.Add($tb_console)
$Form1.Controls.Add($label8)
$Form1.Controls.Add($tb_twilioNum)
$Form1.Controls.Add($tb_twilioToken)
$Form1.Controls.Add($tb_twilioSID)
$Form1.Controls.Add($tb_status2)
$Form1.Controls.Add($tb_status1)
$Form1.Controls.Add($tb_ateraAPI)
$Form1.Controls.Add($label7)
$Form1.Controls.Add($label6)
$Form1.Controls.Add($label5)
$Form1.Controls.Add($table_tech)
$Form1.Controls.Add($label4)
$Form1.Controls.Add($numericUpDown1)
$Form1.Controls.Add($label3)
$Form1.Controls.Add($label2)
$Form1.Controls.Add($label1)
$Form1.Controls.Add($button_save)
$Form1.Controls.Add($button_pause)
$Form1.Controls.Add($button_start)
$Form1.Name = "Form1"
$Form1.Text = "Atera - Assigned Ticket Alerts - Version $Version"




### LET'S GO ###
FirstRun
LoadConfig
if ($AutoStart -eq 1){
    AutoRun
}
$Form1.ShowDialog()
