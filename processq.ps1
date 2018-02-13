#Author: Travis Wirth
#this script will process the open.txt queue, and then move them from the file into the others.

$day = Get-Date -format FileDate
$file = "C:\temp\open.txt"
$file2 = "C:\temp\completed$($day).txt"
$file3 = "C:\temp\errored$($day).txt"
$fileloc = "C:\temp\tickets\"
$recentinclist = "C:\temp\recent.xml"
$openinclist = "C:\temp\open.xml"
$openincsobj = @()
$global:TheList = @()
$recentincsobj1 = @()
$global:RiseStoreList = @()




function getqq {

#clears the open incs list with each new pull
$global:openincsobj = @()


$ie = New-Object -ComObject "InternetExplorer.Application"
$ie.visible=$false
$ie.silent=$true
#echo "Attempting to pull data from site .."
$ie.navigate("https://safeway.service-now.com/incident_list.do?sysparm_query=active%3Dtrue%5Eassigned_toANYTHING%5Eassignment_group%3D43e9cfc86f9b2100f93056df8e3ee4da%5EstateIN1%2C2%2C3")
while($ie.Busy) { Start-Sleep -Milliseconds 50 }
#echo "Extracting relevent data ..."

# K:\SHARED\tickets will be our location for shared tickets

#$list = $ie.Document.IHTMLDocument2_activeElement.outerText  |fl

$list = $ie.Document.frames.document.IHTMLDocument3_documentElement.outerText

$ie.Quit()

$group = "Midrange Operations"
$origin = "Integration"
$origin1 = "BMC Integration"
$origin2 = "TrueSight Integration"

$lines = [regex]::matches( "$($list.ToString())", "Select record for actionPreview(\w+?)\r*?\n*?.*?$($origin)(.+)$($group)").value
$list | clip

$global:testlist = $list
Write-Host "got lines mastah"

$goodlines = @()

$i = 0

foreach ($line in $lines) {

$lines[$i] = $lines[$i] -replace "Select record for actionPreview", ""
$lines[$i] = $lines[$i] -replace "$group", ""
$lines[$i] = $lines[$i] -replace "$origin1", " "
$lines[$i] = $lines[$i] -replace "$origin2", " "
$lines[$i] = $lines[$i] -replace "\r", ""
$lines[$i] = $lines[$i] -replace "\n", ""

$parts = $lines[$i].Split()

$server = $parts[3] -replace "Integration", ""

if ($server -match "patrol") {

    $server = $parts[4]
    } #end if

if ($server -match "Sitescope") {

    $server = $parts[5]
    } #end if

if ($server -match "oem") {

    $server = $parts[4]
    } #end if

if ($server -match "tandem_event") {

    $server = $parts[4]
    } #end if

if ($server -match "alarm") {

    $server = $parts[4]
    } #end if

if ($server -match "performance") {

    $server = $parts[5]
    } #end if


$server = $server -replace ":",""
$server = $server -replace "\n",""
$server = $server -replace "\r",""




$server = $server.ToLower()

#echo " server: $($server)"

# the priority level getting stuck on end of inc number, its annoying me, and confusing
$priority = [regex]::matches( "$($parts[0].ToString())", "\d$").value
$parts[0] = [regex]::matches( "$($parts[0].ToString())", "^\w{10}").value

#echo "inc $($parts[0])"
#echo "priority $priority"

$inc = $parts[0]

$lines[$i] = $parts
$Date = Get-Date
$errortype = ""

switch -wildcard ($lines[$i]) {

"*EAS2013*" {$errortype = "EAS2013"; break}
"*EAS2028*" {$errortype = "EAS2028"; break}
"*EAS2029*" {$errortype = "EAS2029"; break}
"*EAS2391*" {$errortype = "EAS2391"; break}
"*EAS3271*" {$errortype = "EAS3271"; break}
"*EAS3995*" {$errortype = "EAS3995"; break}
"*EAS3997*" {$errortype = "EAS3997"; break}
"*EAS2321*" {$errortype = "EAS2321"; break}
"*RXEAS0032*" {$errortype = "RXEAS0032"; break}


default {$errortype = "Other"}

} #endswitch


#creates a ticket object, that we will then save to file
$ticket = New-Object -TypeName PSOBject
    $ticket | Add-Member -MemberType NoteProperty -Name 'Incident' -Value $inc
    $ticket | Add-Member -MemberType NoteProperty -Name 'Date' -Value $Date
    $ticket | Add-Member -MemberType NoteProperty -Name 'Server' -Value $server
    $ticket | Add-Member -MemberType NoteProperty -Name 'SDesc' -Value $lines[$i]
    $ticket | Add-Member -MemberType NoteProperty -Name 'Priority' -Value $priority
    $ticket | Add-Member -MemberType NoteProperty -Name 'ErrorType' -Value $errortype
    $ticket | Add-Member -MemberType NoteProperty -Name 'Notes' -Value ""

if ($Checksavefile.Checked) {
#creates a folder for the server if one does not exist already
if  ((Test-Path "$fileloc$server") -eq $false) {

md "$fileloc$server"

} #endif
} #endif

if ($Checksavefile.Checked) {
#exports object into xml file
$ticket | select Incident, Date, Server, SDesc, Priority, Notes | Export-Clixml "C:\temp\$inc.xml"
}

#$ticket | Export-Clixml "$fileloc$server\$inc.xml"



$global:openincsobj += $ticket

$i++

} #end foreach

$recentincsobj1 += $global:openincsobj

if ($Checksavefile.Checked) {
$global:openincsobj | Export-Clixml $openinclist
}

} #end getqq


function getrecent {

getqq



if ($Checksavefile.Checked) {
$recentincsobj1 += Import-Clixml $recentinclist
}



$recentincsobj1 += $openincsobj

$recentincsobj1 = $recentincsobj1 | Sort-Object Incident -Unique -descending

$global:TheList = $recentincsobj1

$listBox1.Items.Clear();

foreach ($i in $recentincsobj1){

    $listBox1.items.add("$($i.Incident)             $($i.ErrorType)")
    #$listBox1.Topmost = $True
}


}


function getopen {

getqq

$global:oi = $global:openincsobj


$global:oi = $global:oi | Sort-Object Incident -Unique -descending

$global:TheList = $global:oi

$listBox1.Items.Clear();
foreach ($i in $global:oi)
{
    $listBox1.items.add("$($i.Incident)             $($i.ErrorType)")
}


}


function updatelist {


foreach ($item in $ListBox1.Items) {

if ($item -match $TextSearch.text) {

$ListBox1.SelectedItem = $item
break


} #end if


} #end foreach


} # end updatelist function

# high cpu usage tickets
function checkcpu {



$server = $TheList[$cn].Server
write-host "checkcpu started $server "

$Textnotes.Text +="Checking ..."


#Borrowed the CPU portion from Steve campbell's testy script, as mine seemed inconsistant, and lazy

$ProcNumber = 20

#its only used here, so might as well delcare it here
function RemoveIdle {
	param (
	$processlist
	)
	$newlist = @()
	foreach($p in $processlist){
		if(($p.InstanceName -ne "idle") -and ($p.InstanceName -ne "_total")) { $newlist += $p }
	}
	return $newlist
} #end removeidle function

    #check cpu once now, record it, and write value to a file on a remote host, work in progress

		$CpuCores = (Get-WMIObject Win32_ComputerSystem -ComputerName $server -ErrorAction SilentlyContinue).NumberOfLogicalProcessors 
		if ($CpuCores -lt 1) {
			$CpuInfor = Get-WmiObject Win32_processor -ComputerName $server -ErrorAction SilentlyContinue| select *
			$CpuCores = $CpuInfor.length
		} # end cpu cores if
		# if we can't find the CPU cores then we will assume one for now
		if ($CpuCores -lt 1) { $CpuCores = 1 }
		
		# Debug stub for when I had trouble finding the core amounts
		# Write-Host $CpuCores
		# PressAnyKey

    try {
    
		    $Samples = (Get-Counter "\Process(*)\% Processor Time" -ComputerName $server -ErrorAction silentlycontinue).CounterSamples
		    $SOutput = $Samples | Select `
		    InstanceName,
		    @{Name="CPU";Expression={[Decimal]::Round(($_.CookedValue / $CpuCores), 2)}}
		
		    foreach($OutputLine in $SOutput){
			    if($OutputLine.InstanceName -eq "idle"){
				    $IdleCPU = $OutputLine.CPU
                } #endif
            } #end foreach

            $SOutput = RemoveIdle $SOutput
		    #Write-Host $SOutput
		    #PressAnyKey
		    if($ShowAllProcs){
			    $SOutput = $SOutput | Sort-Object -Descending -Property "CPU"
		    }
            else {
			    $SOutput = $SOutput | Sort-Object -Descending -Property "CPU" | Select-Object -First $ProcNumber
		    } #end else

            $newoutput = @()

             $newoutput += "
             Time of check: $(Get-date -f T)
             Total CPU usage for $($server): $(100 - $IdleCPU)"
		     $newoutput += $SOutput | Select-Object | ft @{E="		"},InstanceName,@{L='CPU %';E={($_.CPU/100).toString('P')}} -AutoSize
  
            $newoutput += echo $cpu

            #since we automating this, no need for clipping for the moment
            #$newoutput | clip.exe
        #echo "bottom of try cpu block, letgs see whats parts we have"
        #echo "$parts"
        #echo $newoutput
           
           $Textnotes.Text = $newoutput | Out-String
           write-host "cpu end successfully"

       } #end try
       catch {
 
       $Textnotes.Text += $error[0] 
       write-host "cpu end errored"

       } #end catch
  
} #end checkcpu function

function checkdiskspace {

# example: phvnprsi: EAS2029 - Logical disk E: free space is at 10.00%


$server = $Thelist[$cn].Server
$parts = $Thelist[$cn].SDesc


# find which logical disk needs to be checked
$DiQ = [regex]::matches( "$parts", "Logical disk \w").value
$DiQ = [regex]::matches( "$DiQ", "\w$").value


$disk = Get-WmiObject Win32_LogicalDisk -ComputerName $server -Filter "DeviceID='$($DiQ):'" |
Select-Object Size,FreeSpace

$ds = $disk.Size  / 1GB
$df = $disk.FreeSpace / 1GB

$dp = $df / $ds

 $ds = "{0:N2}" -f $ds
 $df = "{0:N2}" -f $df 
 $dp = "{0:P2}" -f $dp


 if ($ds -gt 1) {

$Textnotes.Text = "Drive $DiQ details:
Total disk space:    $ds GB
Free disk space:     $df GB
Percentage available $dp"
} #endif
else {

$Textnotes.Text = "$($error[0])"


} #end else


} # end diskspace function

function checkservice {



$server = $Thelist[$cn].Server
$parts = $Thelist[$cn].SDesc.split()

#needs to get the servicename out of setnence

#phcmqz02: EAS2013 - Service SepMasterService is down

#remove is down if it has that, otherwise need to remove trailing periods if it does not show entire name

#echo "parts2 parts1 $parts[-2] $parts[-1]"

if ($parts[-2] -eq "is" -and $parts[-1] -eq "down" ) {
#this will only occur if the name was short enough to not cut off

# iw ill check args to count how man words appear, then taker all but the last 2

    $servicecap = $parts.Length -2
    }
    else {

    $servicecap = $parts.Length
    }


    $service = ""

    $iii = 0
    $iiii = 0
    while ($iiii -lt $servicecap) {

    if ($iii -eq 1) { $service = $service + " " + $parts[$iiii]}

        if ($parts[$iiii] -eq "Service") { 
            $iii = 1
        } #endif
        $iiii++

    } #end while

    $service = $service -replace "^ ", ""


    write-host "Server: $server service name: $service."



    $theerror = 0


    $i = $service.split().length
    write-host "i $i"

    $serviceresults = ""

    $displayname = 0

    #going to try service anem at full length, if ti has 0 results, try with dispaly name instead of name, and if fails, lessen length of name by 1 word and try again

    while ($serviceresults.length -eq 0 -and $i -gt 0) {

    $serviceresults = Get-Service -ComputerName $server -name "$($service)*"

    
    write-host "service: $service."
    write-host "service results:"
    $serviceresults

    if ( $serviceresults.length -eq 0 ) {
        $serviceresults = Get-Service -ComputerName $server -displayname "$($service)*"

        if ( $serviceresults.length -gt 0 ) { $displayname = 1}

    }

    #if we find no results, remove the last word formt he endof the service to try a shortened version
    if ( $serviceresults.length -eq 0 ) {
        $service = $service -replace " \w+$", ""
    }

    $i--

    } #endwhile



    

    
    write-host " final service $service"
    if ( $serviceresults.length -eq 1 ) { 

        if ($displayname -eq 1 ) {
            
            write-host " found result" $serviceresults 
            $output1 = Get-Service  -ComputerName $server -displayname "$($service)*" | select -ExpandProperty status 

            Get-Service -ComputerName $server -displayname "$($service)*" | restart-service 

            $output3 = Get-Service  -ComputerName $server -displayname "$($service)*"  | select -ExpandProperty status

        } #end if
        else {
    
            write-host " found result" $serviceresults 
            $output1 = Get-Service  -ComputerName $server -name "$($service)*" | select -ExpandProperty status 

            Get-Service -ComputerName $server -name "$($service)*" | restart-service

            $output3 = Get-Service  -ComputerName $server -name "$($service)*"  | select -ExpandProperty status 
        } #end else

        $out0 =  "Servicename: $serviceresults.servicename"
        $out1 = "Status of Service:: $service :: before attempting to start: $output1"
        $out3 = "Status of Service:: $service :: after attempting to start: $output3"

        $output2 = "Performed service restart."

        $parts | Out-File $file2 -Append 
        $Textnotes.text += "$out1`r`n$output2`r`n$out3"
        $output = "$out1`r`n$output2`r`n$out3"
        write-host "`r`n$output"


    } #endif

    if ( $serviceresults.length -gt 1 ) { write-host "Will need to look into service list further. Found multiple results" $serviceresults }
    if ( $i -eq 0 ) { write-host "Found no results for service name." }





} #end service function

function pingretailups {

$server = $TheList[$cn].Server

$storenumber = [regex]::matches( "$($server)", '\d{4,6}').value

write-host $storenumber
 

if (test-connection $server -quiet) { 

    $textnotes.text += "$Server responds to pings successfully" 

    $serverstatus = test-connection $server -count 2 | Out-String

    $textnotes.text += $serverstatus

}
else {

    $textnotes.text += "$Server fails to ping"

    if (test-connection -quiet "x$($storenumber))") {

        $textNotes.text += "The x$($storenumber) pings successfully. Sending to Field services to investigate"
    }
    else {
        $textNotes.text += "The x$($storenumber) fails to ping as well. Checking with networking team if there is an issue with the store's network."
    }
} #end of else

} #end function

function RXEAS0032 {

 $textnotes.text = "Checking ..."

$server = $Thelist[$cn].Server
$parts = $Thelist[$cn].SDesc

$storenumber = [regex]::matches( "$($server)", '\d{4,6}').value

$v1 = "b$($storenumber)v1"
$v2 = "b$($storenumber)v2"
$n1 = "b$($storenumber)n1"
$n2 = "b$($storenumber)n2"

# the sdesc doesnt say which node it cant mount to, so lets check both



write-host $storenumber
 

if (test-connection $n1 -quiet -count 1) { 

    $textnotes.text = "$n1 is UP" 
    $n1 = 1
}
else {
    $textnotes.text += "$n1 is DOWN"
    $n1 = 0
} #end of else

if (test-connection $n2 -quiet -count 1) { 

    $textnotes.text += "`r`n$n2 is UP" 
    $n2 = 1
}
else {
    $textnotes.text += "`r`n$n2 is DOWN"
    $n2 = 0
} #end of else

if (test-connection $v1 -quiet -count 1) { 

    $textnotes.text += "`r`n$v1 is UP" 
    $v1 = 1
}
else {
    $textnotes.text += "`r`n$v1 is DOWN"
    $v1 = 0
} #end of else

if (test-connection $v2 -quiet -count 1) { 

    $textnotes.text += "`r`n$v2 is UP"
    $v2 = 1 
}
else {
    $textnotes.text += "`r`n$v2 is DOWN"
    $v2 = 0
} #end of else


# if they both work off n's and they both up, problem solved
if ($n1 -eq 1 -and $n2 -eq 1 ) {
$textnotes.text += "`r`n"
$textnotes.text += "Both nodes are currently up, closing ticket."
}
else {

if ($n1 -eq 1 -or $v1 -eq 1 ) {

if ($n2 -eq 1 -or $v2 -eq 1 ) {
$textnotes.text += "`r`n"
$textnotes.text += "One of each type is up. pharmacy level 2 should be notified to update their script based on the available server types due to RISE. This is assuming that the node in question for this ticket is still down."
} #end if
} #end if


} #end else


if ($n1 -eq 0 -and $v1 -eq 0 ) {
$textnotes.text += "`r`n"
$textnotes.text += "Both n1 and v1 are down, verify which is used for this store and restart as needed."
}

if ($n2 -eq 0 -and $v2 -eq 0 ) {
$textnotes.text += "`r`n"
$textnotes.text += "Both n2 and v2 are down, verify which is used for this store and restart as needed."
}


if ($n1 -eq 0 -and $v1 -eq 0 -and $n2 -eq 0 -and $v2 -eq 0 ) {
$textnotes.text += "`r`n"
$textnotes.text += "All 4 servers are down, likely a power or network issue. Will troubleshoot as needed."
}





} #end function RXEAS0032

function service {

#this one will be used for when we already know the service name

$service = $args.split()
$server = $Thelist[$cn].Server

$i = $service.split().length
    write-host "i $i"

    $serviceresults = ""

    $displayname = 0

    #going to try service anem at full length, if ti has 0 results, try with dispaly name instead of name, and if fails, lessen length of name by 1 word and try again

    while ($serviceresults.length -eq 0 -and $i -gt 0) {

    $serviceresults = Get-Service -ComputerName $server -name "$($service)*"

    write-host "service: $service."
    write-host "service results:"
    $serviceresults

    if ( $serviceresults.length -eq 0 ) {
        $serviceresults = Get-Service -ComputerName $server -displayname "$($service)*"

        if ( $serviceresults.length -gt 0 ) { $displayname = 1}

    }

    #if we find no results, remove the last word formt he endof the service to try a shortened version
    if ( $serviceresults.length -eq 0 ) {
        $service = $service -replace " \w+$", ""
    }

    $i--

    } #endwhile



  write-host " final service $service"
    if ( $serviceresults.length -eq 1 ) { 

        if ($displayname -eq 1 ) {
            
            write-host " found result" $serviceresults 
            $output1 = Get-Service  -ComputerName $server -displayname "$($service)*" | select -ExpandProperty status 

            Get-Service -ComputerName $server -displayname "$($service)*" | restart-service 

            $output3 = Get-Service  -ComputerName $server -displayname "$($service)*"  | select -ExpandProperty status

        } #end if
        else {
    
            write-host " found result" $serviceresults 
            $output1 = Get-Service  -ComputerName $server -name "$($service)*" | select -ExpandProperty status 

            Get-Service -ComputerName $server -name "$($service)*" | restart-service

            $output3 = Get-Service  -ComputerName $server -name "$($service)*"  | select -ExpandProperty status 
        } #end else

        $out0 =  "Servicename: $serviceresults.servicename"
        $out1 = "Status of Service:: $service :: before attempting to start: $output1"
        $out3 = "Status of Service:: $service :: after attempting to start: $output3"

        $output2 = "Performed service restart."

        $parts | Out-File $file2 -Append 
        $Textnotes.text += "$out1`r`n$output2`r`n$out3"
        $output = "$out1`r`n$output2`r`n$out3"
        write-host "`r`n$output"


    } #endif

    if ( $serviceresults.length -gt 1 ) { write-host "Will need to look into service list further. Found multiple results" $serviceresults }
    if ( $i -eq 0 ) { write-host "Found no results for service name." }


} #end function service

function EAS2391 {

# example:
# x291626: EAS2391 - Nastel-MQ Local Manager is down or not responding. LM_Monitor@PROD29\\PROD29\\X291626\\State=Unknown.

$server = $TheList[$cn].Server
$textnotes.text = "checking..."


if (test-connection $server -quiet -count 2) {
# if server pings, then its mq at fault

$textnotes.text = "$server is pinging, try remanaging the node in MQ explorer.
http://osgd-prod.global.safeway.com/"



}
else {
# if server doesnt ping, then server is down, and possibly the ndoe or store as well

$storenumber = [regex]::matches( "$($server)", '\d{4,6}').value

$v1 = "b$($storenumber)v1"
$v2 = "b$($storenumber)v2"
$v3 = "b$($storenumber)v3"
$n1 = "b$($storenumber)n1"
$n2 = "b$($storenumber)n2"


$textnotes.text = "Server not pinging, checking the possible nodes"

if (test-connection $n1 -quiet -count 1) { 

    $textnotes.text += "`r`n$n1 is UP" 
    $n1 = 1
}
else {
    $textnotes.text += "`r`n$n1 is DOWN"
    $n1 = 0
} #end of else

if (test-connection $n2 -quiet -count 1) { 

    $textnotes.text += "`r`n$n2 is UP" 
    $n2 = 1
}
else {
    $textnotes.text += "`r`n$n2 is DOWN"
    $n2 = 0
} #end of else

if (test-connection $v1 -quiet -count 1) { 

    $textnotes.text += "`r`n$v1 is UP" 
    $v1 = 1
}
else {
    $textnotes.text += "`r`n$v1 is DOWN"
    $v1 = 0
} #end of else

if (test-connection $v2 -quiet -count 1) { 

    $textnotes.text += "`r`n$v2 is UP"
    $v2 = 1 
}
else {
    $textnotes.text += "`r`n$v2 is DOWN"
    $v2 = 0
} #end of else

if (test-connection $v3 -quiet -count 1) { 

    $textnotes.text += "`r`n$v3 is UP"
    $v3 = 1 
}
else {
    $textnotes.text += "`r`n$v3 is DOWN"
    $v3 = 0
} #end of else


if ($n1 -eq 0 -and $v1 -eq 0 ) {
$textnotes.text += "`r`n"
$textnotes.text += "Both n1 and v1 are down, verify which is used for this store and restart as needed."
}

if ($n2 -eq 0 -and $v2 -eq 0 ) {
$textnotes.text += "`r`n"
$textnotes.text += "Both n2 and v2 are down, verify which is used for this store and restart as needed."
}


if ($n1 -eq 0 -and $v1 -eq 0 -and $n2 -eq 0 -and $v2 -eq 0 ) {
$textnotes.text += "`r`n"
$textnotes.text += "All 4 servers are down, likely a power or network issue. Will troubleshoot as needed."
}



} #end else


} #end function





function fixit {



# first paramater will be the server
#second param will be errortype
# third and onward will be the additional details


$inctype = $null

# if the args come in as just 1 long argument, we want to break it up !
$parts = $args.Split()

#this should only happen if you forgot to put the short description after invoking this script
if ($parts[0] -eq $null) { Write-Host " no arguments for fixit function given"}



#rewriting into switch statement for easier future expansion

write-host "fixit errortype $($parts[1])"

switch -wildcard ($parts[1]) {


"*EAS2013*" {checkservice}
"*EAS2028*" {checkcpu}
"*EAS2029*" {checkdiskspace}
"*EAS2391*" {EAS2391}
"*EAS3271*" {pingretailups}
"*EAS3975*" {service "patrol*"}
"*EAS3995*" {service "patrol*"}
"*EAS3977*" {service "patrol*"}
"*EAS2321*" {service "patrol*"}
"*RXEAS0032*" {RXEAS0032}
"*Other*" { $textNotes.text = "Can't fix this one automatically."}




} #endswitch


$global:TheList[$cn].Notes += $Textnotes.text


#needs to verify they are windows servers?


<# this can be used to check if server matches windows standard naming
if ( $server -match '^.{3}[n,m]') {

    echo "Server matches windows naming..."

    } #end windows check if
    
    #>







} #end fixit function





#rise tool functions

function findrisestore {


$TextRISEMac.text = "Checking ..."


$store = $ListboxRISE.SelectedItem

if ($store.length -eq 4) {

    $SampleServer = "b$($store)v2"

    write-host $SampleServer

    $hostname = [System.Net.Dns]::GetHostbyname("$SampleServer")

    $hostname = $hostname.HostName

    $fullstorename = [regex]::Matches("$hostname", "^\w{7}")
    $fullstorename = [regex]::Matches("$fullstorename", "\w{6}$")

    $store = $fullstorename

} #end if

if ($store -match "\d{6}") {
$fullstorename = $store
}



$pserver = "p$($fullstorename)"



$pnetad = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $pserver


write-host "hostname $fullstorename"


$macaddress = @()

foreach ($mac in $pnetad) {

if ($mac.ipaddress -ne $null ) {
$macaddress += $mac.macaddress
}

} # end foreach


if ($macaddress.length -gt 1 ) {
$TextRISEMac.text =  "$macaddress"


}

if ($macaddress.length -eq 1 ) {
$TextRISEMac.text =  "$macaddress"

}

if ($macaddress.length -lt 1 ) {

$TextRISEMac.text =  "No mac addresses found with an IP for $pserver. Verify you are connected to VPN, and $pserver server is pingable."
}

$global:RiseStoreList[$ListboxRISE.SelectedIndex].FullStore = $fullstorename
$global:RiseStoreList[$ListboxRISE.SelectedIndex].Mac = $TextRISEMac.text

write-host " text $($TextRISEMac.text)"
write-host " risestorelsit mac $($global:RiseStoreList[$ListboxRISE.SelectedIndex].Mac)"
write-host " selected item $($ListboxRISE.SelectedIndex)"

} # end function risestore


function RISEstep17 {

$textRISEStep17.text = "Checking ..."


$store = $ListboxRISE.SelectedItem

if ($store.length -eq 4) {

    if ( $global:RiseStoreList[$ListboxRISE.SelectedIndex].FullStore -match "\d{6}" ) {

        $fullstorename = $global:RiseStoreList[$ListboxRISE.SelectedIndex].FullStore
    }
    else {

        $SampleServer = "b$($store)v2"
        write-host $SampleServer
        $hostname = [System.Net.Dns]::GetHostbyname("$SampleServer")
        $hostname = $hostname.HostName
        $fullstorename = [regex]::Matches("$hostname", "^\w{7}")
        $fullstorename = [regex]::Matches("$fullstorename", "\w{6}$")

        $store = $fullstorename
    } #end else
} #end if

if ($store -match "\d{6}") {
$fullstorename = $store
}

# up until now we jsut making sure we have the full store number stored within $fullstorename
#now we pull the network config for the v2 and v3, and compare against what we expect to see

$v1 = "b$($fullstorename)v1.safeway.com" # not used yet in code, this was made prior to R2 rise
$v2 = "b$($fullstorename)v2.safeway.com"
$v3 = "b$($fullstorename)v3.safeway.com"

$ramlist = @()






. "C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1"
Connect-VIServer -server $textvcenter.text


write-host " vserver $vserver"

Get-VMHost | Where-Object {$_.Name -eq $v2 -or $_.Name -eq $v3} |
%{$ramlist += $_.Name; $ramlist += $_.MemoryTotalMB; Get-View $_.ID} |
%{$esxname = $_.Name; Get-View $_.ConfigManager.NetworkSystem} |
%{ foreach($physnic in $_.NetworkInfo.Pnic){
    $pnicInfo = $_.QueryNetworkHint($physnic.Device)
    foreach($hint in $pnicInfo){
      Write-Host $esxname $physnic.Device
      
      $physnic.Device
      if( $hint.ConnectedSwitchPort ) {
        write-host "$($hint.ConnectedSwitchPort | out-string)"
        
        $switchport = $hint.ConnectedSwitchPort.DevId
        $port = $hint.ConnectedSwitchPort.PortId

      

        $switchport = [regex]::matches($switchport, "^.{6}").value
        
        $switchport = [regex]::matches($switchport, ".$").value
     
        $port = [regex]::matches($port, "\d*$").value

      
        write-host "device $($physnic.Device)"
        
        switch -wildcard ($physnic.Device) {

        "*vmnic0*" {
            if ( $esxname -match $v2 ) {

                if ($switchport -match "x" -and $port -eq "7") {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic0 = "Good"
                }
                else {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic0 = "$v2 Switch X Port7 expected, but found Switch $switchport Port $port instead."
                }

            } #end if
             if ( $esxname -eq $v3 ) {

                if ($switchport -eq "x" -and $port -eq "9") {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic0 = "Good"
                }
                else {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic0 = "$v3 Switch X Port9 expected, but found Switch $switchport Port $port instead."
                }

            } #end if
        }
        "*vmnic1*" {
          if ( $esxname -eq $v2 ) {

                if ($switchport -eq "x" -and $port -eq "13") {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic1 = "Good"
                }
                else {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic1 = "$v2 Switch X Port13 expected, but found Switch $switchport Port $port instead."
                }

            } #end if
             if ( $esxname -eq $v3 ) {

                if ($switchport -eq "x" -and $port -eq "15") {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic1 = "Good"
                }
                else {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic1 = "$v3 Switch X Port15 expected, but found Switch $switchport Port $port instead."
                }

            } #end if
        }
        "*vmnic2*" {
          if ( $esxname -eq $v2 ) {

                if ($switchport -eq "y" -and $port -eq "7") {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic2 = "Good"
                }
                else {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic2 = "$v2 switch Y Port7 expected, but found Switch $switchport Port $port instead."
                }

            } #end if
             if ( $esxname -eq $v3 ) {

                if ($switchport -eq "y" -and $port -eq "9") {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic2 = "Good"
                }
                else {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic2 = "$v3 Switch Y Port9 expected, but found Switch $switchport Port $port instead."
                }

            } #end if
        }
        "*vmnic3*" {
                    if ( $esxname -eq $v2 ) {

                if ($switchport -eq "y" -and $port -eq "13") {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic3 = "Good"
                }
                else {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic3 = "$v2 Switch Y Port13 expected, but found Switch $switchport Port $port instead."
                }

            } #end if
             if ( $esxname -eq $v3 ) {

                if ($switchport -eq "y" -and $port -eq "15") {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic3 = "Good"
                }
                else {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic3 = "$v3 Switch Y Port15 expected, but found Switch $switchport Port $port instead."
                }

            } #end if
        }

        "*vmnic4*" {
                    if ( $esxname -eq $v2 ) {

                if ($switchport -eq "x" -and $port -eq "8") {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic4 = "Good"
                }
                else {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic4 = "$v2 Switch X Port8 expected, but found Switch $switchport Port $port instead."
                }

            } #end if
             if ( $esxname -eq $v3 ) {

                if ($switchport -eq "x" -and $port -eq "10") {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic4 = "Good"
                }
                else {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic4 = "$v3 Switch X Port10 expected, but found Switch $switchport Port $port instead."
                }

            } #end if
        }
        "*vmnic5*" {
                    if ( $esxname -eq $v2 ) {

                if ($switchport -eq "y" -and $port -eq "8") {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic5 = "Good"
                }
                else {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic5 = "$v2 Switch Y Port8 expected, but found Switch $switchport Port $port instead."
                }

            } #end if
             if ( $esxname -eq $v3 ) {

                if ($switchport -eq "y" -and $port -eq "10") {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic5 = "Good"
                }
                else {
                $global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic5 = "$v3 Switch Y Port10 expected, but found Switch $switchport Port $port instead."
                }

            } #end if
        }



        } #end switch
      } # end if
      else {
        Write-Host "$($physnic.Device | out-string) No CDP information available."; Write-Host
      } #end else

    } # end foreach
  }#end foreach
} #end of percent

#} #end foreach


#sort RAM


write-host $ramlist

# formats numbers to look like GB, but kinda not really
$ramlist[1] = $ramlist[1] / 1000
$ramlist[1] = "{0:N1}" -f $ramlist[1]
$ramlist[3] = $ramlist[3] / 1000
$ramlist[3] = "{0:N1}" -f $ramlist[3]

if ($ramlist[0] -eq $v2 ) {
$global:RiseStoreList[$ListboxRISE.SelectedIndex].v2ram = $ramlist[1]
}
if ($ramlist[0] -eq $v3 ) {
$global:RiseStoreList[$ListboxRISE.SelectedIndex].v3ram = $ramlist[1]
}

if ($ramlist[2] -eq $v2 ) {
$global:RiseStoreList[$ListboxRISE.SelectedIndex].v2ram = $ramlist[3] 
}
if ($ramlist[2] -eq $v3 ) {
$global:RiseStoreList[$ListboxRISE.SelectedIndex].v3ram = $ramlist[3]
}



# for sorting out nics
$step17reply = $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic0, $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic1, $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic2, $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic3, $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic4, $global:RiseStoreList[$ListboxRISE.SelectedIndex].v2vmnic5,$global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic0,$global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic1,$global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic2,$global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic3,$global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic4,$global:RiseStoreList[$ListboxRISE.SelectedIndex].v3vmnic5


$textRISEStep17.text = ""

foreach ( $vmnic in $step17reply) {

if ($vmnic -ne "Good") {

$textRISEStep17.text += $vmnic

}


} #end foreach

$risestep17counter = 0
if ($textRISEStep17.text -eq "") {
$risestep17counter++
$textRISEStep17.text = "All NIC's are good."
}

#paste RAM info at bottom


if ($global:RiseStoreList[$ListboxRISE.SelectedIndex].v2ram -gt 32 -and $global:RiseStoreList[$ListboxRISE.SelectedIndex].v3ram -gt 32) {
$risestep17counter++
$textRISEStep17.text += "`r`n"
$textRISEStep17.text += "RAM is good on both servers."
}

if ($global:RiseStoreList[$ListboxRISE.SelectedIndex].v2ram -lt 32) {
$textRISEStep17.text += "`r`n"
$textRISEStep17.text += "v2 RAM: $($global:RiseStoreList[$ListboxRISE.SelectedIndex].v2ram)GB, expected to see 32 GB"

}

if ($global:RiseStoreList[$ListboxRISE.SelectedIndex].v3ram -lt 32) {
$textRISEStep17.text += "`r`n"
$textRISEStep17.text += "v3 RAM: $($global:RiseStoreList[$ListboxRISE.SelectedIndex].v3ram)GB, expected to see 32 GB"

}

# this only occurs if both step went smoothly
if ($risestep17counter -eq 2) {
$textRISEStep17.text += "`r`n"
$textRISEStep17.text += "Step 17 completed successfully."

}

$global:RiseStoreList[$ListboxRISE.SelectedIndex].Step17text = $textRISEStep17.text



} #end function



function addRISEstore {


$addedstores = @()

$addedstores = $TextRISEadd.text.split()


foreach ($addedstore in $addedstores) {

$newstore = New-Object -TypeName PSOBject
    $newstore | Add-Member -MemberType NoteProperty -Name 'Store' -Value $addedStore
    $newstore | Add-Member -MemberType NoteProperty -Name 'FullStore' -Value ""
    $newstore | Add-Member -MemberType NoteProperty -Name 'Mac' -Value "Not yet checked."
    $newstore | Add-Member -MemberType NoteProperty -Name 'v2vmnic0' -Value "untested"
    $newstore | Add-Member -MemberType NoteProperty -Name 'v2vmnic1' -Value "untested"
    $newstore | Add-Member -MemberType NoteProperty -Name 'v2vmnic2' -Value "untested"
    $newstore | Add-Member -MemberType NoteProperty -Name 'v2vmnic3' -Value "untested"
    $newstore | Add-Member -MemberType NoteProperty -Name 'v2vmnic4' -Value "untested"
    $newstore | Add-Member -MemberType NoteProperty -Name 'v2vmnic5' -Value "untested"
    $newstore | Add-Member -MemberType NoteProperty -Name 'v3vmnic0' -Value "untested"
    $newstore | Add-Member -MemberType NoteProperty -Name 'v3vmnic1' -Value "untested"
    $newstore | Add-Member -MemberType NoteProperty -Name 'v3vmnic2' -Value "untested"
    $newstore | Add-Member -MemberType NoteProperty -Name 'v3vmnic3' -Value "untested"
    $newstore | Add-Member -MemberType NoteProperty -Name 'v3vmnic4' -Value "untested"
    $newstore | Add-Member -MemberType NoteProperty -Name 'v3vmnic5' -Value "untested"
    $newstore | Add-Member -MemberType NoteProperty -Name 'v2ram' -Value "untested"
    $newstore | Add-Member -MemberType NoteProperty -Name 'v3ram' -Value "untested"
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step1' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step2' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step3' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step4' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step5' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step6' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step7' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step8' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step9' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step10' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step11' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step12' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step13' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step14' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step15' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step16' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step17' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step18' -Value $false
    $newstore | Add-Member -MemberType NoteProperty -Name 'Step17text' -Value "Not yet checked"


    



    $listboxRISE.items.add($addedstore)


    $global:RiseStoreList += $newstore


    } # end foreach





} # end function addstore











#Form creation section


#region GUI assemblies
#Load Assemblies

[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null

$net = New-Object -ComObject Wscript.Network

 
 

#Draw background form

$Form = New-Object System.Windows.Forms.Form

 $Form.width = 800

 $Form.height = 600

 $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable

 $Form.Text = "Incident managment"

 #$Form.maximumsize = New-Object System.Drawing.Size(525,350)

 $Form.startposition = "centerscreen"

 $Form.KeyPreview = $True

 $Form.Add_KeyDown({if ($_.KeyCode -eq "Enter") {}})

 $Form.Add_KeyDown({if ($_.KeyCode -eq "Escape")

     {$Form.Close()}})

#endregion




#region incs tab gui

     #itmes for incs tab

#create listbox on left
#this form will show inc #'s, and server name,

$global:ListBox1 = New-Object System.Windows.Forms.ListBox

$ListBox1.Location = New-Object System.Drawing.Size(90,75)

$ListBox1.Size = New-Object System.Drawing.Size(260,20)

$ListBox1.Height = 460

$listbox1.ScrollAlwaysVisible = $True

#$listBox1.TabIndex = -1




#how do i word this ?!?!!?
#$ListBox1.SelectionMode = $listbox.selectionmode.one




#create textboxes
$TextServer = New-Object System.Windows.Forms.textbox

$TextServer.Location = New-Object System.Drawing.Size(560,20)

$TextServer.Size = New-Object System.Drawing.Size(80,20)

$TextServer.multiline = $true



$TextInc = New-Object System.Windows.Forms.textbox

$TextInc.Location = New-Object System.Drawing.Size(440,20)

$TextInc.Size = New-Object System.Drawing.Size(80,20)

$TextInc.multiline = $true



$Textpriority = New-Object System.Windows.Forms.textbox

$Textpriority.Location = New-Object System.Drawing.Size(660,20)

$Textpriority.Size = New-Object System.Drawing.Size(20,20)

$Textpriority.multiline = $true



$Textsdesc = New-Object System.Windows.Forms.textbox

$Textsdesc.Location = New-Object System.Drawing.Size(440,40)

$Textsdesc.Size = New-Object System.Drawing.Size(260,60)

$Textsdesc.multiline = $true



$Textnotes = New-Object System.Windows.Forms.textbox

$Textnotes.Location = New-Object System.Drawing.Size(440,100)

$Textnotes.Size = New-Object System.Drawing.Size(260,400)

$Textnotes.multiline = $true


$TextSearch = New-Object System.Windows.Forms.textbox

$TextSearch.Location = New-Object System.Drawing.Size(90,50)

$TextSearch.Size = New-Object System.Drawing.Size(260,20)

$TextSearch.multiline = $true 

#Create buttons

 $fixbutton = new-object System.Windows.Forms.Button

 $fixbutton.Location = new-object System.Drawing.Size(705,100)

 $fixbutton.Size = new-object System.Drawing.Size(80,20)

 $fixbutton.Text = "Fix it"

 $fixbutton.Add_Click({fixit $global:TheList[$cn].server $global:TheList[$cn].errortype})



 $updatelist = new-object System.Windows.Forms.Button

 $updatelist.Location = new-object System.Drawing.Size(5,50)

 $updatelist.Size = new-object System.Drawing.Size(80,20)

 $updatelist.Text = "Search"

 $updatelist.Add_Click({updatelist})



 $CopyInc = new-object System.Windows.Forms.Button

 $CopyInc.Location = new-object System.Drawing.Size(419,20)

 $CopyInc.Size = new-object System.Drawing.Size(20,20)

 $CopyInc.Text = "Copy Inc"

 $CopyInc.Add_Click({$TextInc.Text | clip.exe})


 $CopyServer = new-object System.Windows.Forms.Button

 $CopyServer.Location = new-object System.Drawing.Size(539,20)

 $CopyServer.Size = new-object System.Drawing.Size(20,20)

 $CopyServer.Text = "C"

 $CopyServer.Add_Click({$TextServer.Text | clip.exe})


 $CopyNotes = new-object System.Windows.Forms.Button

 $CopyNotes.Location = new-object System.Drawing.Size(419,100)

 $CopyNotes.Size = new-object System.Drawing.Size(20,20)

 $CopyNotes.Text = "C"

 $CopyNotes.Add_Click({$TextNotes.Text | clip.exe})


 $CopySDesc = new-object System.Windows.Forms.Button

 $CopySDesc.Location = new-object System.Drawing.Size(419,40)

 $CopySDesc.Size = new-object System.Drawing.Size(20,20)

 $CopySDesc.Text = "C"

 $CopySDesc.Add_Click({$TextSDesc.Text | clip.exe})


 $Openincs = new-object System.Windows.Forms.Button

 $Openincs.Location = new-object System.Drawing.Size(5,130)

 $Openincs.Size = new-object System.Drawing.Size(80,20)

 $Openincs.Text = "Open"

 $Openincs.Add_Click({getopen})


 $RecentIncs = new-object System.Windows.Forms.Button

 $RecentIncs.Location = new-object System.Drawing.Size(5,155)

 $RecentIncs.Size = new-object System.Drawing.Size(80,20)

 $RecentIncs.Text = "Recent"

 $RecentIncs.Add_Click({getrecent})



 # create checkbox

 
 $Checksavefile = new-object System.windows.forms.Checkbox

 $Checksavefile.Location = new-object System.Drawing.Size(5,180)

 $Checksavefile.Size = new-object System.Drawing.Size(20,20)

 $checksavefile.text = "Save?"

 $checksavefile.checked = $false




 #create Labels

$LabelSaveFile = New-Object System.Windows.Forms.Label

$LabelSaveFile.Text = "Save?"

$LabelSaveFile.AutoSize = $True

$LabelSaveFile.Location = new-object System.Drawing.Size(25,182)

#endregion

#region GUI rise page


#create listbox

$ListBoxRISE = New-Object System.Windows.Forms.ListBox

$ListBoxRISE.Location = New-Object System.Drawing.Size(90,75)

$ListBoxRISE.Size = New-Object System.Drawing.Size(260,20)

$ListBoxRISE.Height = 460

$ListBoxRISE.ScrollAlwaysVisible = $True


#create buttons

$buttonRISEMac = new-object System.Windows.Forms.Button

 $buttonRISEMac.Location = new-object System.Drawing.Size(705,50)

 $buttonRISEMac.Size = new-object System.Drawing.Size(80,20)

 $buttonRISEMac.Text = "Step 9.4"

 $buttonRISEMac.Add_Click({findrisestore})


  $ButtonRISEAdd = new-object System.Windows.Forms.Button

 $ButtonRISEAdd.Location = new-object System.Drawing.Size(5,50)

 $ButtonRISEAdd.Size = new-object System.Drawing.Size(80,20)

 $ButtonRISEAdd.Text = "Add Store"

 $ButtonRISEAdd.Add_Click({addRISEStore})


  $ButtonCopyMac = new-object System.Windows.Forms.Button

 $ButtonCopyMac.Location = new-object System.Drawing.Size(419,40)

 $ButtonCopyMac.Size = new-object System.Drawing.Size(20,20)

 $ButtonCopyMac.Text = "C"

 $ButtonCopyMac.Add_Click({$TextRISEMac.Text | clip.exe})

 
 $ButtonStep17 = new-object System.Windows.Forms.Button

 $ButtonStep17.Location = new-object System.Drawing.Size(705,100)

 $ButtonStep17.Size = new-object System.Drawing.Size(80,20)

 $ButtonStep17.Text = "Step 17"

 $ButtonStep17.Add_Click({RISEstep17})

 #region create check boxes and corresponding labels


  $CheckStep1 = new-object System.windows.forms.Checkbox

 $CheckStep1.Location = new-object System.Drawing.Size(440,310)

 $CheckStep1.Size = new-object System.Drawing.Size(20,20)

 $CheckStep1.text = "Step 1"

 $CheckStep1.checked = $false

 $CheckStep1.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step1 = $CheckStep1.Checked })


 $LabelStep1 = New-Object System.Windows.Forms.Label

$LabelStep1.Text = "Step1"

$LabelStep1.AutoSize = $True

$LabelStep1.Location = new-object System.Drawing.Size(460,313)


  $CheckStep2 = new-object System.windows.forms.Checkbox

 $CheckStep2.Location = new-object System.Drawing.Size(440,330)

 $CheckStep2.Size = new-object System.Drawing.Size(20,20)

 $CheckStep2.text = "Step 2"

 $CheckStep2.checked = $false

 $CheckStep2.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step2 = $CheckStep2.Checked })


  $LabelStep2 = New-Object System.Windows.Forms.Label

$LabelStep2.Text = "Step2"

$LabelStep2.AutoSize = $True

$LabelStep2.Location = new-object System.Drawing.Size(460,333)



  $CheckStep3 = new-object System.windows.forms.Checkbox

 $CheckStep3.Location = new-object System.Drawing.Size(440,350)

 $CheckStep3.Size = new-object System.Drawing.Size(20,20)

 $CheckStep3.text = "Save?"

 $CheckStep3.checked = $false

 $CheckStep3.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step3 = $CheckStep3.Checked })


  $LabelStep3 = New-Object System.Windows.Forms.Label

$LabelStep3.Text = "Step3"

$LabelStep3.AutoSize = $True

$LabelStep3.Location = new-object System.Drawing.Size(460,353)


 
 $CheckStep4 = new-object System.windows.forms.Checkbox

 $CheckStep4.Location = new-object System.Drawing.Size(440,370)

 $CheckStep4.Size = new-object System.Drawing.Size(20,20)

 $CheckStep4.text = "Save?"

 $CheckStep4.checked = $false

 $CheckStep4.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step4 = $CheckStep4.Checked })


  $LabelStep4 = New-Object System.Windows.Forms.Label

$LabelStep4.Text = "Step4"

$LabelStep4.AutoSize = $True

$LabelStep4.Location = new-object System.Drawing.Size(460,373)




  $CheckStep5 = new-object System.windows.forms.Checkbox

 $CheckStep5.Location = new-object System.Drawing.Size(440,390)

 $CheckStep5.Size = new-object System.Drawing.Size(20,20)

 $CheckStep5.text = "Save?"

 $CheckStep5.checked = $false

 $CheckStep5.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step5 = $CheckStep5.Checked })


  $LabelStep5 = New-Object System.Windows.Forms.Label

$LabelStep5.Text = "Step5"

$LabelStep5.AutoSize = $True

$LabelStep5.Location = new-object System.Drawing.Size(460,393)



  $CheckStep6 = new-object System.windows.forms.Checkbox

 $CheckStep6.Location = new-object System.Drawing.Size(440,410)

 $CheckStep6.Size = new-object System.Drawing.Size(20,20)

 $CheckStep6.text = "Save?"

 $CheckStep6.checked = $false

 $CheckStep6.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step6 = $CheckStep6.Checked })


  $LabelStep6 = New-Object System.Windows.Forms.Label

$LabelStep6.Text = "Step6"

$LabelStep6.AutoSize = $True

$LabelStep6.Location = new-object System.Drawing.Size(460,413)


  $CheckStep7 = new-object System.windows.forms.Checkbox

 $CheckStep7.Location = new-object System.Drawing.Size(440,430)

 $CheckStep7.Size = new-object System.Drawing.Size(20,20)

 $CheckStep7.text = "Save?"

 $CheckStep7.checked = $false

 $CheckStep7.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step7 = $CheckStep7.Checked })


  $LabelStep7 = New-Object System.Windows.Forms.Label

$LabelStep7.Text = "Step7"

$LabelStep7.AutoSize = $True

$LabelStep7.Location = new-object System.Drawing.Size(460,433)


  $CheckStep8 = new-object System.windows.forms.Checkbox

 $CheckStep8.Location = new-object System.Drawing.Size(440,450)

 $CheckStep8.Size = new-object System.Drawing.Size(20,20)

 $CheckStep8.text = "Save?"

 $CheckStep8.checked = $false

 $CheckStep8.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step8 = $CheckStep8.Checked })


  $LabelStep8 = New-Object System.Windows.Forms.Label

$LabelStep8.Text = "Step8"

$LabelStep8.AutoSize = $True

$LabelStep8.Location = new-object System.Drawing.Size(460,453)


  $CheckStep9 = new-object System.windows.forms.Checkbox

 $CheckStep9.Location = new-object System.Drawing.Size(440,470)

 $CheckStep9.Size = new-object System.Drawing.Size(20,20)

 $CheckStep9.text = "Step9"

 $CheckStep9.checked = $false

 $CheckStep9.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step9 = $CheckStep9.Checked })


  $LabelStep9 = New-Object System.Windows.Forms.Label

$LabelStep9.Text = "Step9"

$LabelStep9.AutoSize = $True

$LabelStep9.Location = new-object System.Drawing.Size(460,473)


  $CheckStep10 = new-object System.windows.forms.Checkbox

 $CheckStep10.Location = new-object System.Drawing.Size(440,490)

 $CheckStep10.Size = new-object System.Drawing.Size(20,20)

 $CheckStep10.text = "Save?"

 $CheckStep10.checked = $false

 $CheckStep10.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step10 = $CheckStep10.Checked })


  $LabelStep10 = New-Object System.Windows.Forms.Label

$LabelStep10.Text = "Step10"

$LabelStep10.AutoSize = $True

$LabelStep10.Location = new-object System.Drawing.Size(460,493)



  $CheckStep11 = new-object System.windows.forms.Checkbox

 $CheckStep11.Location = new-object System.Drawing.Size(540,310)

 $CheckStep11.Size = new-object System.Drawing.Size(20,20)

 $CheckStep11.text = "Save?"

 $CheckStep11.checked = $false

 $CheckStep11.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step11 = $CheckStep11.Checked })


  $LabelStep11 = New-Object System.Windows.Forms.Label

$LabelStep11.Text = "Step11"

$LabelStep11.AutoSize = $True

$LabelStep11.Location = new-object System.Drawing.Size(560,313)


  $CheckStep12 = new-object System.windows.forms.Checkbox

 $CheckStep12.Location = new-object System.Drawing.Size(540,330)

 $CheckStep12.Size = new-object System.Drawing.Size(20,20)

 $CheckStep12.text = "Save?"

 $CheckStep12.checked = $false

 $CheckStep12.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step12 = $CheckStep12.Checked })


  $LabelStep12 = New-Object System.Windows.Forms.Label

$LabelStep12.Text = "Step12"

$LabelStep12.AutoSize = $True

$LabelStep12.Location = new-object System.Drawing.Size(560,333)


  $CheckStep13 = new-object System.windows.forms.Checkbox

 $CheckStep13.Location = new-object System.Drawing.Size(540,350)

 $CheckStep13.Size = new-object System.Drawing.Size(20,20)

 $CheckStep13.text = "Save?"

 $CheckStep13.checked = $false

 $CheckStep13.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step13 = $CheckStep13.Checked })


  $LabelStep13 = New-Object System.Windows.Forms.Label

$LabelStep13.Text = "Step13"

$LabelStep13.AutoSize = $True

$LabelStep13.Location = new-object System.Drawing.Size(560,353)


  $CheckStep14 = new-object System.windows.forms.Checkbox

 $CheckStep14.Location = new-object System.Drawing.Size(540,370)

 $CheckStep14.Size = new-object System.Drawing.Size(20,20)

 $CheckStep14.text = "Save?"

 $CheckStep14.checked = $false

 $CheckStep14.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step14 = $CheckStep14.Checked })


  $LabelStep14 = New-Object System.Windows.Forms.Label

$LabelStep14.Text = "Step14"

$LabelStep14.AutoSize = $True

$LabelStep14.Location = new-object System.Drawing.Size(560,373)


 $CheckStep15 = new-object System.windows.forms.Checkbox

 $CheckStep15.Location = new-object System.Drawing.Size(540,390)

 $CheckStep15.Size = new-object System.Drawing.Size(20,20)

 $CheckStep15.text = "Save?"

 $CheckStep15.checked = $false

 $CheckStep15.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step15 = $CheckStep15.Checked })


  $LabelStep15 = New-Object System.Windows.Forms.Label

$LabelStep15.Text = "Step15"

$LabelStep15.AutoSize = $True

$LabelStep15.Location = new-object System.Drawing.Size(560,393)



  $CheckStep16 = new-object System.windows.forms.Checkbox

 $CheckStep16.Location = new-object System.Drawing.Size(540,410)

 $CheckStep16.Size = new-object System.Drawing.Size(20,20)

 $CheckStep16.text = "Save?"

 $CheckStep16.checked = $false

 $CheckStep16.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step16 = $CheckStep16.Checked })


  $LabelStep16 = New-Object System.Windows.Forms.Label

$LabelStep16.Text = "Step16"

$LabelStep16.AutoSize = $True

$LabelStep16.Location = new-object System.Drawing.Size(560,413)


  $CheckStep17 = new-object System.windows.forms.Checkbox

 $CheckStep17.Location = new-object System.Drawing.Size(540,430)

 $CheckStep17.Size = new-object System.Drawing.Size(20,20)

 $CheckStep17.text = "Save?"

 $CheckStep17.checked = $false

 $CheckStep17.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step17 = $CheckStep17.Checked })


  $LabelStep17 = New-Object System.Windows.Forms.Label

$LabelStep17.Text = "Step17"

$LabelStep17.AutoSize = $True

$LabelStep17.Location = new-object System.Drawing.Size(560,433)


  $CheckStep18 = new-object System.windows.forms.Checkbox

 $CheckStep18.Location = new-object System.Drawing.Size(540,450)

 $CheckStep18.Size = new-object System.Drawing.Size(20,20)

 $CheckStep18.text = "Save?"

 $CheckStep18.checked = $false

 $CheckStep18.Add_CheckStateChanged({ $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step18 = $CheckStep18.Checked })



 $LabelStep18 = New-Object System.Windows.Forms.Label
 
$LabelStep18.Text = "Step18"

$LabelStep18.AutoSize = $True

$LabelStep18.Location = new-object System.Drawing.Size(560,453)

#endregion 


 #create textboxes



 
 $textvcenter = New-Object System.Windows.Forms.textbox

$textvcenter.Location = New-Object System.Drawing.Size(90,20)

$textvcenter.Size = New-Object System.Drawing.Size(260,20)

$textvcenter.text = "pn01046d"


$LabelVcenter = New-Object System.Windows.Forms.Label
 
$LabelVcenter.Text = "VCenter:"

$LabelVcenter.AutoSize = $True

$LabelVcenter.Location = new-object System.Drawing.Size(40,20)


 $TextRISEMac = New-Object System.Windows.Forms.textbox

$TextRISEMac.Location = New-Object System.Drawing.Size(440,40)

$TextRISEMac.Size = New-Object System.Drawing.Size(260,60)

$TextRISEMac.multiline = $true


$TextRISEAdd = New-Object System.Windows.Forms.textbox

$TextRISEAdd.Location = New-Object System.Drawing.Size(90,50)

$TextRISEAdd.Size = New-Object System.Drawing.Size(260,20)

$TextRISEAdd.multiline = $true 


$textRISEStep17 = New-Object System.Windows.Forms.textbox

$textRISEStep17.Location = New-Object System.Drawing.Size(440,100)

$textRISEStep17.Size = New-Object System.Drawing.Size(260,200)

$textRISEStep17.multiline = $true




#endregion

#region create tabpages and tabcontrol

#tools page
#rise page

# unsure how this works exactly
#$System_Drawing_Point.X = 5
#$System_Drawing_Point.Y = 5

$TabControl = New-object System.Windows.Forms.TabControl
$TabInc = New-Object System.windows.Forms.Tabpage
$TabTools = New-Object System.windows.Forms.Tabpage
$TabRISE = New-Object System.windows.Forms.Tabpage

$TabControl.DataBindings.DefaultDataSourceUpdateMode = 0
$TabControl.Location = new-object System.Drawing.Size(0,0)
$TabControl.Name = "TabControl"
$TabControl.size = new-object System.Drawing.Size(800,600)


$TabInc.Name = "Incs"
$TabInc.Text = "Incs"


$TabTools.Name = "Tools"
$TabTools.Text = "Tools"


$TabRISE.Name = "RISE"
$TabRISE.Text = "RISE"

#endregion



 # update text boxes absed on selected item in listbox



$listBox1.add_SelectedIndexChanged({


$global:cn = $global:listbox1.SelectedIndex

#write-host "SelectedIndex= "$listbox1.SelectedIndex
#write-host 'SelectedItem= ' $listbox1.SelectedItem

$Textsdesc.text = $global:TheList[$cn].sdesc
$Textinc.text = $global:TheList[$cn].Incident
$textserver.text = $global:TheList[$cn].server
$Textpriority.text = $global:TheList[$cn].priority
$Textnotes.text = $global:TheList[$cn].Notes

})



#update on selected for rise tab

$listBoxRISE.add_SelectedIndexChanged({


$global:cn = $global:listboxRISE.SelectedIndex

#write-host "SelectedIndex= "$listbox1.SelectedIndex
#write-host 'SelectedItem= ' $listbox1.SelectedItem

$TextRISEMac.text = $global:RiseStoreList[$listboxRISE.SelectedIndex].Mac

$CheckStep1.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step1
$CheckStep2.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step2
$CheckStep3.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step3
$CheckStep4.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step4
$CheckStep5.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step5
$CheckStep6.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step6
$CheckStep7.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step7
$CheckStep8.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step8
$CheckStep9.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step9
$CheckStep10.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step10
$CheckStep11.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step11
$CheckStep12.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step12
$CheckStep13.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step13
$CheckStep14.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step14
$CheckStep15.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step15
$CheckStep16.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step16
$CheckStep17.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step17
$CheckStep18.Checked = $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step18

$textRISEStep17.text =  $global:RiseStoreList[$ListboxRISE.SelectedIndex].Step17text

}) # end index changed


 








 #add tabs

 $Form.Controls.Add($TabControl)
 $TabControl.Controls.Add($Tabinc)
 $TabControl.Controls.Add($TabRISE)
 $TabControl.Controls.Add($Tabtools)



 function tabinc {

 $TabRISE.Controls.Clear()


 #add labels

 $TabInc.Controls.Add($LabelSaveFile)


 #add checkboxes

 $TabInc.Controls.Add($checksavefile)

#add the buttons

$TabInc.Controls.Add($Openincs)
$TabInc.Controls.Add($RecentIncs)
$TabInc.Controls.Add($CopyServer)
$TabInc.Controls.Add($CopyInc)
$TabInc.Controls.Add($CopySDesc)
$TabInc.Controls.Add($CopyNotes)
$TabInc.Controls.Add($updatelist)
$TabInc.Controls.Add($fixbutton)



#add textboxes
$TabInc.Controls.Add($TextServer)
$TabInc.Controls.Add($Textinc)
$TabInc.Controls.Add($Textnotes)
$TabInc.Controls.Add($Textpriority)
$TabInc.Controls.Add($Textsdesc)
$TabInc.Controls.Add($TextSearch)

#add the listbox
$TabInc.Controls.Add($ListBox1)




} #end inc tab function 



function Tabtools {

$TabRISE.Controls.Clear()
$TabInc.Controls.Clear()


}


function TabRISE {

$TabInc.Controls.Clear()


#add buttons
$TabRISE.Controls.Add($buttonRISEMac)
$TabRISE.Controls.Add($buttonRISEadd)
$TabRISE.Controls.Add($ButtonCopyMac)
$TabRISE.Controls.Add($ButtonStep17)

#add checkboxes
$TabRISE.Controls.Add($CheckStep1)
$TabRISE.Controls.Add($CheckStep2)
$TabRISE.Controls.Add($CheckStep3)
$TabRISE.Controls.Add($CheckStep4)
$TabRISE.Controls.Add($CheckStep5)
$TabRISE.Controls.Add($CheckStep6)
$TabRISE.Controls.Add($CheckStep7)
$TabRISE.Controls.Add($CheckStep8)
$TabRISE.Controls.Add($CheckStep9)
$TabRISE.Controls.Add($CheckStep10)
$TabRISE.Controls.Add($CheckStep11)
$TabRISE.Controls.Add($CheckStep12)
$TabRISE.Controls.Add($CheckStep13)
$TabRISE.Controls.Add($CheckStep14)
$TabRISE.Controls.Add($CheckStep15)
$TabRISE.Controls.Add($CheckStep16)
$TabRISE.Controls.Add($CheckStep17)
$TabRISE.Controls.Add($CheckStep18)

#add labels
$TabRISE.Controls.Add($LabelStep1)
$TabRISE.Controls.Add($LabelStep2)
$TabRISE.Controls.Add($LabelStep3)
$TabRISE.Controls.Add($LabelStep4)
$TabRISE.Controls.Add($LabelStep5)
$TabRISE.Controls.Add($LabelStep6)
$TabRISE.Controls.Add($LabelStep7)
$TabRISE.Controls.Add($LabelStep8)
$TabRISE.Controls.Add($LabelStep9)
$TabRISE.Controls.Add($LabelStep10)
$TabRISE.Controls.Add($LabelStep11)
$TabRISE.Controls.Add($LabelStep12)
$TabRISE.Controls.Add($LabelStep13)
$TabRISE.Controls.Add($LabelStep14)
$TabRISE.Controls.Add($LabelStep15)
$TabRISE.Controls.Add($LabelStep16)
$TabRISE.Controls.Add($LabelStep17)
$TabRISE.Controls.Add($LabelStep18)

$TabRISE.Controls.Add($LabelVcenter)



#add textbox
$TabRISE.Controls.Add($TextRISEMac)
$TabRISE.Controls.Add($TextRISEAdd)
$TabRISE.Controls.Add($textRISEStep17)
$TabRISE.Controls.Add($textvcenter)



#add listbox
$TabRISE.Controls.Add($ListBoxRISE)

}




function tabhandler {

switch -wildcard ($Tabcontrol.SelectedTab) {

"*Incs*" {TabInc}
"*Tools*" {TabTools}
"*RISE*" {TabRISE}


}



} # end tab handler function




$TabControl.add_SelectedIndexChanged({tabhandler})


#calling tabinc function to load its info once initially.
TabInc

$Form.Add_Shown({$Form.Activate()})
$Form.ShowDialog()




<#  notes:


#Author: Travis Wirth

#you can ignore most of these bottom notes, as some of these are simply ideas, not whats implemented

#for disk space, it should check the diskpace on recent servers that had an issue to see what their normal usage is, maybe once a day or so - currently i do not have it tracking tickets over a long period

#another script to read that file to help form a basis of servers that need either increased space/cpu

#another script for retail boxes to verify whats actually down at a store, and whether its likely a RISE store or not

# needs to verify whether a store is effected by projects, such as rise

#currently I have one to pull data and sort out thr short description. This needs to store them in a list somewhere, so that this script is not run agaisnt the same server twice needlessly, and to store the notes for each ticket. it should not disapear formt he lsit until it can identify that notes have been added tot he ticket.

#it woudl be great if i could find a way to submit work notes automatically as well, but if i look into this first, then I might find a way to submit notes, as well as see fi they exist


#buttons i should add: 
copy inc,
copy server, 
copy both of those+ short description(good for giving to other groups)
copy notes, 
save notes, 
run fix(gray out when unavailable for that ticket), 
check queue, 
turn on or off autochecks or autofix

search archived inc (maybe add later, as this would involve a bit more


# to streamline this, i shoudl update my process handlign code to create an object for each ticket with the mentioned properties

# use import and exportclixml to save data in a more logical fashion, as people will no logner be looking directly at text files
 


#a separate object should keep track of incs and servers, and then as they disapear from open queue, be added to recent property for quick reference

# object names should be named based on inc, however be stored in folders based on servername for easier searching into archives later



#>