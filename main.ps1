<#
.SYNOPSIS
	
.DESCRIPTION
	
.PARAMETER 
	
.EXAMPLE
	
.NOTES
	
.LINK
	
#>

## Load Assemblys
Add-Type -AssemblyName presentationframework
##import JSON config file
$script:config = get-content("$psscriptroot\config.json") | convertfrom-json
##import JSON language file

$language=(Get-WinUserLanguageList).languageTag
$script:language = get-content("$($script:config.languageFolder)$($script:config.messageFile)") | convertfrom-json
## Debugging
$DebugPreference = "Continue"

#region main functions
function backToWork() {
    $flrs = Get-childitem $($script:config.floorFolder) -filter *.png
    $floors = ""
    foreach ($flr in $flrs.name) {
        $flr = $flr.substring(0, $flr.Length - 4)
        $floors = $floors + "<ComboBoxItem>$flr</ComboBoxItem>"
    }
    
    [xml]$xaml = @"
    <Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Back to Work" Height="250" Width="350"
        ResizeMode="NoResize">
    <Grid>
        <DatePicker x:Name="from" HorizontalAlignment="Left" Margin="180,35,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="user" HorizontalAlignment="Left" Margin="180,5,0,0" Text="$(get-friendlyName)" TextWrapping="Wrap" VerticalAlignment="Top" Width="130"/>
        <ComboBox x:Name="floor" HorizontalAlignment="Left" Margin="180,65,0,0" VerticalAlignment="Top" Width="40">
            $floors
        </ComboBox>
        <Label Content="$($script:language.$($script:config.language).msg001)" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" Width="175" />
        <Label Content="$($script:language.$($script:config.language).msg002)" HorizontalAlignment="Left" Margin="10,60,0,0" VerticalAlignment="Top" Width="175"/>
        <Label x:Name="freeSeats" Content="$($script:language.$($script:config.language).msg019)" HorizontalAlignment="Left" Margin="10,90,0,0" VerticalAlignment="Top" Width="175"/>
        <Label Content="$($script:language.$($script:config.language).msg003)" HorizontalAlignment="Left" Margin="10,30,0,0" VerticalAlignment="Top" Width="175" />
        <Label x:Name="Error" Content="Label" HorizontalAlignment="Center" Margin="0,175,0,0" VerticalAlignment="Top" Width="330" Background="#FFFFD2D2" Foreground="Red" FontWeight="Bold" Visibility="Hidden" />
        <Button x:Name="search" Content="$($script:language.$($script:config.language).msg004)" HorizontalAlignment="Left" Margin="120,130,0,0" VerticalAlignment="Top"/>
        <Button x:Name="view" Content="$($script:language.$($script:config.language).msg018)" HorizontalAlignment="Left" Margin="180,130,0,0" VerticalAlignment="Top"/>
    </Grid>
</Window>

"@

    $CM = @{}
    $CM.ContextMenu = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))
    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object -Process {
        $CM.$($_.Name) = $CM.ContextMenu.FindName($_.Name)
    }
	
    $cm.floor.add_SelectionChanged({
        $fromdate = Get-date($cm.from.SelectedDate) -Format $script:config.dateFormat
        $floor = $cm.floor.SelectedItem.content
        if (!($fromdate)) {
            $cm.Error.content = "$($script:language.$($script:config.language).msg006)"
            $cm.error.visibility = "Visible"
        } else {
            $cm.freeSeats.content= "$($script:language.$($script:config.language).msg019) $(get-freeseats $floor $fromdate) "
        }
    })
    $cm.from.add_SelectedDateChanged({
        $fromdate = Get-date($cm.from.SelectedDate) -Format $script:config.dateFormat
        $floor = $cm.floor.SelectedItem.content
        if ($floor -eq "") {
            $cm.Error.content = "$($script:language.$($script:config.language).msg009)"
            $cm.error.visibility = "Visible"
        } else {
            $cm.freeSeats.content= "$($script:language.$($script:config.language).msg019) $(get-freeseats $floor $fromdate) "
        }
    })

    $cm.view.add_click({
        $user = $cm.user.text
        $fromdate = Get-date($cm.from.SelectedDate) -Format $script:config.dateFormat
        $floor = $cm.floor.text
        if (!($fromdate)) {
            $cm.Error.content = "$($script:language.$($script:config.language).msg006)"
            $cm.error.visibility = "Visible"
        } else {
            selectwindow $user $floor $fromdate -view
        }
    })
    
    $cm.search.Add_Click({
        $user = $cm.user.text
        $fromdate = Get-date($cm.from.SelectedDate) -Format $script:config.dateFormat
        $floor = $cm.floor.text
        $result, $seat, $bfloor = check-userbooking $fromdate $user
    
        If ($result) {
            $cm.Error.content = "$($script:language.$($script:config.language).msg005): $bfloor $($seat.replace('s',''))"
            $cm.error.visibility = "Visible"
        }
        elseif (!($fromdate)) {
            $cm.Error.content = "$($script:language.$($script:config.language).msg006)"
            $cm.error.visibility = "Visible"

        }
        elseif ((get-date $fromdate) -le (get-date)) {
            $cm.Error.content = "$($script:language.$($script:config.language).msg007)"
            $cm.error.visibility = "Visible"

        }elseif ($floor -eq "") {
            $cm.Error.content = "$($script:language.$($script:config.language).msg009)"
            $cm.error.visibility = "Visible"

        } else {
            $seats, $count = get-seats $floor "$fromdate"
            $totcount = 0
            $content = get-content("$($script:config.bookingsFolder)$($script:config.floorsFile)") | convertfrom-json
            foreach ($item in $($content."$($floor)").psobject.properties) {
                $totcount += 1
            }
            $totcount
            $Percent = ($count * 100) / $totcount
            $maxres = $script:config.maxAllocationPercent

            if ($percent -ge $maxres) {
                $cm.Error.content = "$($script:language.$($script:config.language).msg010) > $maxres %"
                $cm.error.visibility = "Visible"
            }  else {

                $CM.ContextMenu.Close()
                selectwindow $user $floor $fromdate
            }
        }
    })   
    $CM.ContextMenu.ShowDialog()
}

function selectwindow($user, $floor, $fromdate,[switch]$view) {
    
    $png = New-Object System.Drawing.Bitmap "$($script:config.floorFolder)$($floor).png"
    $height = $($png.Height) + 40
    $width = $($png.Width) + 20
    if($view){
        $seats, $count = get-seats $floor $fromdate -view
    } else {
        $seats, $count = get-seats $floor $fromdate
    }
    
    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
       Title="Back to Work" Height="$height" Width="$width"
       ResizeMode="NoResize">
    <Border x:Name="MainMenuBorder" >
        <Grid >
            <Grid.Background>
                <ImageBrush ImageSource="$($script:config.floorFolder)$($floor).png"/>
            </Grid.Background>
            $seats
        </Grid>
    </Border>
</Window>
"@
    # Create the Window and get all elements
    $CM = @{}
    $CM.ContextMenu = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))
    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object -Process {
        $CM.$($_.Name) = $CM.ContextMenu.FindName($_.Name)
    }
    #handler for dynamic MenuItems
    [System.Windows.RoutedEventHandler]$clickHandler = {
        $seat = $($_.OriginalSource.Name)
        $CM.ContextMenu.Close()
        set-booking $seat $floor $fromdate $user
    }  
    $cm.MainMenuBorder.AddHandler([System.Windows.Controls.Button]::ClickEvent, $clickHandler)
         
    $CM.ContextMenu.ShowDialog()
}

Function set-booking($seat, $floor, $fromdate, $user) {
    $snumber = $seat.Replace("s", "")
    [xml]$xaml = @"
    <Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
       Title="Back to Work" Height="250" Width="350"
       ResizeMode="NoResize">
    <Grid>
        <TextBox x:Name="from" HorizontalAlignment="Left" Text="$fromdate" Margin="175,35,0,0" VerticalAlignment="Top" IsEnabled="False"/>
        <TextBox x:Name="user" HorizontalAlignment="Left" Text="$user" Margin="175,0,0,0" VerticalAlignment="Top" Width="125" IsEnabled="False"/>
        <TextBox x:Name="floor" HorizontalAlignment="Left" Text="$floor" Margin="175,105,0,0" VerticalAlignment="Top" Width="40" IsEnabled="False"/>
        <TextBox x:Name="seat" HorizontalAlignment="Left" Text="$snumber" Margin="175,70,0,0" VerticalAlignment="Top" Width="40" IsEnabled="False"/>
        <Label Content="$($script:language.$($script:config.language).msg001)" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" />
        <Label Content="$($script:language.$($script:config.language).msg002)" HorizontalAlignment="Left" Margin="10,100,0,0" VerticalAlignment="Top"/>
        <Label Content="$($script:language.$($script:config.language).msg011)" HorizontalAlignment="Left" Margin="10,60,0,0" VerticalAlignment="Top"/>
        <Label Content="$($script:language.$($script:config.language).msg003)" HorizontalAlignment="Left" Margin="10,30,0,0" VerticalAlignment="Top" />
        <Label x:Name="Error" Content="Label" HorizontalAlignment="Center" Margin="0,135,0,0" VerticalAlignment="Top" Width="340" Background="#FFFFD2D2" Foreground="Red" FontWeight="Bold" Visibility="Hidden" />

        <Button x:Name="save" Content="$($script:language.$($script:config.language).msg012)" HorizontalAlignment="Left" Margin="85,165,0,0" VerticalAlignment="Top"/>
        <Button x:Name="cancel" Content="$($script:language.$($script:config.language).msg013)" HorizontalAlignment="Left" Margin="175,165,0,0" VerticalAlignment="Top"/>
        <Button x:Name="neu" Content="$($script:language.$($script:config.language).msg014)" HorizontalAlignment="Left" Margin="90,165,0,0" VerticalAlignment="Top" Visibility="hidden"/>
    </Grid>

</Window>

"@
    $CM = @{}
    $CM.ContextMenu = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))
    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object -Process {
        $CM.$($_.Name) = $CM.ContextMenu.FindName($_.Name)
    }

    $cm.save.Add_Click({
            $user = $cm.user.text
            $fromdate= $cm.from.text
            $floor = $cm.floor.text
            
            Save $floor $seat $fromdate $user
            $cm.Error.content = "$($script:language.$($script:config.language).msg015)"
            $cm.error.visibility = "Visible"
            $cm.error.background = "#FFD2FFD6"
            $cm.error.foreground = "green"
            $cm.save.visibility = "hidden"
            $cm.cancel.content = "$($script:language.$($script:config.language).msg016)"
            $cm.neu.Visibility = "Visible"
                    
        })

    $cm.neu.Add_Click({
            $CM.ContextMenu.Close()
            backToWork
        })

    $cm.cancel.Add_Click({
            $CM.ContextMenu.Close()
        })

    $CM.ContextMenu.ShowDialog()
}
#endregion main functions

#region helper functions
function get-seats($floor, $fromdate,[switch]$view) {
    $buttons = @()
    $content = get-content("$($script:config.bookingsFolder)$($script:config.floorsFile)") | convertfrom-json
    $count = 0
    foreach ($item in $($content."$($floor)").psobject.properties) {
        $seat = $item.name
        $h, $w = $item.value.split(",")
        $h = $h - 11
        $w = $w - 11
        $booked, $user = get-bookings $floor $seat $fromdate 
        
        if($view){
            $IsEnabled="False"  
            
        } else{
            $IsEnabled="True"
            
        }
        #$booked
        if ($booked) {
            $count += 1
            $color = "red"
            $buttons = $buttons + @"
    <TextBlock x:Name="$($item.name)" ToolTip="$($script:language.$($script:config.language).msg017) $user" HorizontalAlignment="Left" Margin="$h,$w,0,0" Text="$($item.name.replace("s",""))" TextWrapping="Wrap" VerticalAlignment="Top" Background="Red" Height="22" Width="22" TextAlignment="Center"/>
"@
        }
        else {
            $color = "lime"
            $buttons = $buttons + @"
    <Button x:Name="$($item.name)" Content="$($item.name.replace("s",""))" HorizontalAlignment="Left" Margin="$h,$w,0,0" Height="22" Width="22" VerticalAlignment="Top" Background="$color"  IsEnabled="$IsEnabled"/>
"@
        }
    }
    return $buttons , $count
}

function get-bookings($number, $seat, $fromdate) {
    $content = get-content("$($script:config.bookingsFolder)$($script:config.bookingsFile)") | convertfrom-json
    $bookings = $content.$number.$seat
    if ($bookings.count -gt 0) {
        $bookings | ForEach-Object {
            if (!($booking)) {
                $date =$_.date
                if ($fromdate -eq $date) {
                    $booking = $true
                    $user = $_.user
                }
                else {
                    $booking = $false
                }
            }
        }
    }
    else {
        $booking = $false
    }
    return $booking, $user
}

function get-DateArray {
    Param(
        [Parameter(Mandatory = $True)]
        [DateTime]$StartDate,
        [Parameter(Mandatory = $True)]
        [DateTime]$EndDate
    )

    [Array]$DateArray = @()
    while ((Get-Date $StartDate.tostring($script:config.dateFormat)) -lt (Get-Date $EndDate.tostring($script:config.dateFormat))) {
        $NextDate = Get-Date $StartDate 
        $DateArray += $NextDate
        $startDate = $startDate.AddDays(1)
    }
    $DateArray += $EndDate 

    Return $DateArray
}

function Save($floor, $seat, $date, $user) {    
    $content = get-content("$($script:config.bookingsFolder)$($script:config.bookingsFile)") | convertfrom-json
    $hash = [pscustomobject]@{date = "$date"; user = "$user" }
    $content.$floor.$seat += $hash
    $content.$floor.$seat
    $write = $false
    while (!($write)) {
        Try {
            $now = get-date -format yyyyMMdd_HHmmss
            if (!(Test-Path  "$($script:config.bookingsBackupFolder)")) {
                New-Item -path "$($script:config.bookingsBackupFolder)" -type Directory
            }    
            copy-item "$($script:config.bookingsFolder)$($script:config.bookingsFile)" "$($script:config.bookingsBackupFolder)\bookings_$now.json"
            $content | ConvertTo-Json -Depth 3 | out-file "$($script:config.bookingsFolder)$($script:config.bookingsFile)"
            $Write = $true
        }
        catch {
            $write = $false
        }
    }
}

function check-userbooking($date, $user) {
    $result = $false
    $content = get-content("$($script:config.bookingsFolder)$($script:config.bookingsFile)") | convertfrom-json
    $floors = get-content("$($script:config.bookingsFolder)$($script:config.floorsFile)") | convertfrom-json
    foreach($floor in $floors.psobject.properties){
        if(!($floor.name -match "^maxAllocationPercent")){
            $floor=$floor.name 
            foreach ($item in $content."$floor".psobject.properties) {
                $str = $content."$floor"."$($item.name)"
                if ($str | Where-Object { ($_."date" -EQ "$date") -and ($_."user" -eq "$user") }) {
                    $result = $true
                    $seat = "$($item.name)"
                }
            }
        }
    }
    return $result, $seat, $floor
}

function get-friendlyName() {
    #because we didnt want to ask ad for a friendly name we query registry. add other source if u want.
    $key = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Common\UserInfo"
    $value = "UserName"
    $result = Get-ItemProperty -Path Registry::$key $value
    return $result.username
}

function get-freeSeats($floor,$fromdate){
    $seats, $count=get-seats $floor $fromdate
    $bookedSeats=$count
    $totcount = 0
    $content = get-content("$($script:config.bookingsFolder)$($script:config.floorsFile)") | convertfrom-json
    foreach ($item in $($content."$($floor)").psobject.properties) {
        $totcount += 1
    }
    if($null -eq $content."maxAllocationPercent$floor"){
        $maxAlloc=$script:config.maxAllocationPercent
    } else {
        $maxAlloc=$content."maxAllocationPercent$floor"
    }
    [decimal]$totcount=[math]::floor(($totcount*$maxAlloc/100))
    $totcount= $totcount-$bookedSeats
    if(($totcount-$bookedSeats) -lt 0 ){
        return 0
    } else {
        return $totcount
    }
}

function get-jsondata(){
 if($script:config.bookingsFolder){
    
 }
 If($true){

 }

}
#endregion

#region main execution
function main() {
    #executions start
    Start-Transcript -Path $script:config.logPath -Append
    #check if app already runs
    if (!(get-process | Where-Object processname -eq "powershell" | select-object * | Where-Object mainWindowTitle -eq "Back to Work")) {
        backToWork
    } else {
        write-host "Application already running"
    }
    Stop-Transcript
    #executions end
}

if ($MyInvocation.InvocationName -ne '.') {
    main
}
#endregion