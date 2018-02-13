# script to balance retail vm's
#Author: Travis Wirth


# lets make an object for each store found in a vcenter, along with what vm's exist for the store

# if the vm exists, it counts the number of vm's
# give each vm a weight for how much resources it uses to help balance - this might come later
# check how many hosts the store has to balance across 

#migrate vm's to match balance


# if test is false, it will actualy perform migrations, otherwise it just looks at data if $true
$test = $true


# command to vmotion: Get-VM VM1 | Move-VM -Destination ( Get-VMHost ESXHost2 )

# max vm per host assuming all 3 hosts are up
$threshold = 2

#max vm per host threshold if theres only 2 hosts
$threshold2 = 3

#name of array of clusters
$cluster = @()

$clustertemplate = New-Object PSObject
        $clustertemplate | Add-Member -type NoteProperty -Name 'name' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'V1name' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'V2name' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'V3name' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'V1status' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'V2status' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'V3status' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'V1vcount' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'V2vcount' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'V3vcount' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'V1RAM' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'V2RAM' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'V3RAM' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'passed' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'Vvirt' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'Fvirt' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'Ivirt' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'Pvirt' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'Qvirt' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'Xvirt' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'GoodHosts' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'GoodV1' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'GoodV2' -Value 0
        $clustertemplate | Add-Member -type NoteProperty -Name 'GoodV3' -Value 0




# command to add more to cluster array: $cluster += $clustertemplate.psobject.copy()



# lets connect to our test cluster

. "C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1"
Connect-VIServer -server phvnprv4



#cool, we have our template, lets pulla list of clusters
[System.Collections.ArrayList]$clusterlist = get-cluster

# im seeing some as -shipped, lets ignore these
$i = 0
foreach ($item in $clusterlist ) {

if ($item -imatch "shipped") {

$clusterlist.Remove($i)
}

$i++
} #foreach


# up next, lets check how many hosts each one has, follow by how many are powered on! followed by RAM check.
# we will also add tot he array at this point, time to start recording data

foreach ($item in $clusterlist ) {

$cluster += $clustertemplate.psobject.copy()

$cluster[-1].Name = $item.Name

$clusterhosts = Get-Cluster $item.name | Get-VMHost


# lets sort so we know which is which

foreach ($chost in $clusterhosts) {


#lets also check what vm's are on each host during this loop

$vtemp = 0
$ftemp = 0
$qtemp = 0
$xtemp = 0
$itemp = 0
$ptemp = 0


$vmlist = $chost | get-vm

#this loops checks if the vm exists on this host, and marks it as 10 temporarily, we do not know if this is v1,v2, or v3, yet.
foreach ($vmitem in $vmlist ) {

    if ($vmitem.Name -imatch "v$($item.name)") {
    $vtemp = 10
    }#if

    if ($vmitem.Name -imatch "f$($item.name)") {
    $ftemp = 10
    }#if

    if ($vmitem.Name -imatch "x$($item.name)") {
    $xtemp = 10
    }#if

    if ($vmitem.Name -imatch "p$($item.name)") {
    $ptemp = 10
    }#if

    if ($vmitem.Name -imatch "i$($item.name)") {
    $itemp = 10
    }#if

    if ($vmitem.Name -imatch "q$($item.name)") {
    $qtemp = 10
    }#if
} #foreach



if ($chost.Name -imatch "v1") {

$cluster[-1].V1Name = $chost.Name
$cluster[-1].V1Status = $chost.ConnectionState
$cluster[-1].V1RAM = $chost.MemoryTotalGB

if ($Itemp -eq 10 ) {
$cluster[-1].Ivirt = 1
}
if ($Qtemp -eq 10 ) {
$cluster[-1].Qvirt = 1
}
if ($Ftemp -eq 10 ) {
$cluster[-1].Fvirt = 1
}
if ($Ptemp -eq 10 ) {
$cluster[-1].Pvirt = 1
}
if ($Xtemp -eq 10 ) {
$cluster[-1].Xvirt = 1
}
if ($Vtemp -eq 10 ) {
$cluster[-1].Vvirt = 1
}

} #if v1 check


#v2 check
if ($chost.Name -imatch "v2") {

$cluster[-1].V2Name = $chost.Name
$cluster[-1].V2Status = $chost.ConnectionState
$cluster[-1].V2RAM = $chost.MemoryTotalGB

if ($Itemp -eq 10 ) {
$cluster[-1].Ivirt = 2
}
if ($Qtemp -eq 10 ) {
$cluster[-1].Qvirt = 2
}
if ($Ftemp -eq 10 ) {
$cluster[-1].Fvirt = 2
}
if ($Ptemp -eq 10 ) {
$cluster[-1].Pvirt = 2
}
if ($Xtemp -eq 10 ) {
$cluster[-1].Xvirt = 2
}
if ($Vtemp -eq 10 ) {
$cluster[-1].Vvirt = 2
}

} #if v2 check


#v3 check
if ($chost.Name -imatch "v3") {

$cluster[-1].V3Name = $chost.Name
$cluster[-1].V3Status = $chost.ConnectionState
$cluster[-1].V3RAM = $chost.MemoryTotalGB

if ($Itemp -eq 10 ) {
$cluster[-1].Ivirt = 3
}
if ($Qtemp -eq 10 ) {
$cluster[-1].Qvirt = 3
}
if ($Ftemp -eq 10 ) {
$cluster[-1].Fvirt = 3
}
if ($Ptemp -eq 10 ) {
$cluster[-1].Pvirt = 3
}
if ($Xtemp -eq 10 ) {
$cluster[-1].Xvirt = 3
}
if ($Vtemp -eq 10 ) {
$cluster[-1].Vvirt = 3
}

} #if v3 check


} #foreach host loop
} #foreach cluster loop






# okay, now we have all info on the current setup, we need to decide how to do this. 
# First, lets ignore any servers with less than optimal RAM, as these are likely going to have issues, or be down for repairs soon. These should have no VM's on them. Maybe make this optional later, or set a threshold.
# This also means, we now need to figure out how many valid servers each cluster has, and which they are.


# weights, the bigger the number, the more processing power it tends to need.
# lets also try and keep Q and I separate for redundancy
# v2 runs D, so lets keep it lighter if possible

#RISe doc shows a default of V2: D, X, Q, I  V3: P, F - seems silly to have q and i on same box imo


# maybe ... v1 f x, v2 d q, v3 p i

#possible weights as well

#$vweight = 7
#$xweight = 7
#$qweight = 7
#$pweight = 10
#$fweight = 10
#$iweight = 10



# so we only need to balance ones that have more than 1 valid host, lets swrite out the possible ways we will balance

# if we see a valid  host with no vm's, lets check if tis valid and pull from the other 1 or 2

# if we see a host with more than 2 vm's lets check if it has 3 valid hosts

# if we see a host with more than 3 vm's lets check if it has 2 valid hosts

# check RAM to ensure its a good host

# i shoudl go abck and verifyt hat the vm is also running to be counted


#this loop is for checking which servers meet our criteria, separated so it can be changed more easily later, also counts how many relevant virtuals each has currently.
$i = 0
foreach ($item in $cluster) {


# if we see more than 31 gigs of ram and showing connected status, lets assume its a usable host.
if ($item.v1ram -gt 31 -and $item.v1status -eq "Connected") {

$cluster[$i].goodv1 = 1
}

if ($item.v2ram -gt 31 -and $item.v2status -eq "Connected") {

$cluster[$i].goodv2 = 1
}

if ($item.v3ram -gt 31 -and $item.v3status -eq "Connected") {

$cluster[$i].goodv3 = 1
}


$cluster[$i].GoodHosts = $cluster[$i].goodv1 + $cluster[$i].goodv2 + $cluster[$i].goodv3

# this counts how many virtuals each host has
if ( $item.ivirt -eq 1 ) {$item.v1vcount++}
if ( $item.xvirt -eq 1 ) {$item.v1vcount++}
if ( $item.pvirt -eq 1 ) {$item.v1vcount++}
if ( $item.fvirt -eq 1 ) {$item.v1vcount++}
if ( $item.vvirt -eq 1 ) {$item.v1vcount++}
if ( $item.qvirt -eq 1 ) {$item.v1vcount++}

if ( $item.ivirt -eq 2 ) {$item.v2vcount++}
if ( $item.xvirt -eq 2 ) {$item.v2vcount++}
if ( $item.pvirt -eq 2 ) {$item.v2vcount++}
if ( $item.fvirt -eq 2 ) {$item.v2vcount++}
if ( $item.vvirt -eq 2 ) {$item.v2vcount++}
if ( $item.qvirt -eq 2 ) {$item.v2vcount++}

if ( $item.ivirt -eq 3 ) {$item.v3vcount++}
if ( $item.xvirt -eq 3 ) {$item.v3vcount++}
if ( $item.pvirt -eq 3 ) {$item.v3vcount++}
if ( $item.fvirt -eq 3 ) {$item.v3vcount++}
if ( $item.vvirt -eq 3 ) {$item.v3vcount++}
if ( $item.qvirt -eq 3 ) {$item.v3vcount++}

$i++
} # foreach



# lets count how many vm's are on each host




#now lets identify what we actually want to do with the vm's based on their setup

#if it has more than threashhold, then take however many spare it has, and put them in pending status
#pending status will not migrate, just put it up for grabs in script for the second loop to distribute

#check others for being too low, and then grab form pending pool if it finds any

#so loop each cluster twice, once tto put some in pending, then again to assign them to spare

write-host " before testgrounds"

#im puting a huge if block here to not actually move stuff without being out of test mode
if ($test -eq $false) {

write-host " entering testgrounds"

$i = 0
foreach ($item in $cluster) { 


# start of v1's check over threshold
if ($item.v1vcount -gt $threshold) {

write-host "$($item.v1name) has $($item.v1vcount), over threadhold of $threshold"

$over = $item.v1vcount - $threshold

# lets mark however many over threshold as pending, we will put pending items as belonging to 9, rather thant he real respective owner

if ($item.xvirt -eq 1 -and $over -gt 0) { 
$item.xvirt = 9
$over = $over - 1
} # if
if ($item.ivirt -eq 1 -and $over -gt 0) { 
$item.ivirt = 9
$over = $over - 1
} # if
if ($item.pvirt -eq 1 -and $over -gt 0) { 
$item.pvirt = 9
$over = $over - 1
} # if
if ($item.fvirt -eq 1 -and $over -gt 0) { 
$item.fvirt = 9
$over = $over - 1
} # if
if ($item.vvirt -eq 1 -and $over -gt 0) { 
$item.vvirt = 9
$over = $over - 1
} # if
if ($item.qvirt -eq 1 -and $over -gt 0) { 
$item.qvirt = 9
$over = $over - 1
} # if

} # if goodv1 check

if ($item.v2vcount -gt $threshold) {

write-host "$($item.v2name) has $($item.v2vcount), over threadhold of $threshold"

$over = $item.v2vcount - $threshold

# lets mark however many over threshold as pending, we will put pending items as belonging to 9

if ($item.xvirt -eq 2 -and $over -gt 0) { 
$item.xvirt = 9
$over = $over - 1
} # if
if ($item.ivirt -eq 2 -and $over -gt 0) { 
$item.ivirt = 9
$over = $over - 1
} # if
if ($item.pvirt -eq 2 -and $over -gt 0) { 
$item.pvirt = 9
$over = $over - 1
} # if
if ($item.fvirt -eq 2 -and $over -gt 0) { 
$item.fvirt = 9
$over = $over - 1
} # if
if ($item.vvirt -eq 2 -and $over -gt 0) { 
$item.vvirt = 9
$over = $over - 1
} # if
if ($item.qvirt -eq 2 -and $over -gt 0) { 
$item.qvirt = 9
$over = $over - 1
} # if

} # if goodv1 check

if ($item.v3vcount -gt $threshold) {

write-host "$($item.v3name) has $($item.v3vcount), over threadhold of $threshold"

$over = $item.v3vcount - $threshold

# lets mark however many over threshold as pending, we will put pending items as belonging to 9

if ($item.xvirt -eq 3 -and $over -gt 0) { 
$item.xvirt = 9
$over = $over - 1
write-host " pending $($item.xvirt)"
} # if
if ($item.ivirt -eq 3 -and $over -gt 0) { 
$item.ivirt = 9
$over = $over - 1
write-host " pending $($item.ivirt)"
} # if
if ($item.pvirt -eq 3 -and $over -gt 0) { 
$item.pvirt = 9
$over = $over - 1
write-host " pending $($item.pvirt)"
} # if
if ($item.fvirt -eq 3 -and $over -gt 0) { 
$item.fvirt = 9
$over = $over - 1
write-host " pending $($item.fvirt)"
} # if
if ($item.vvirt -eq 3 -and $over -gt 0) { 
$item.vvirt = 9
$over = $over - 1
write-host " pending $($item.vvirt)"
} # if
if ($item.qvirt -eq 3 -and $over -gt 0) { 
$item.qvirt = 9
$over = $over - 1
write-host " pending $($item.qvirt)"
} # if

} #if v3 count

# so now we know if any has less thant he threshhold ,and only put as many vm's into pending as needed to resolve it


#v1 set

if ($item.GoodV1 -eq 1) {
$over = $item.v1vcount

if ( $over -lt $threshold -and $item.vvirt -eq 9 ) {
# this is where the magic command happens to vmotion the server over!
write-host " moving v$($item.name) to  $($item.v1name)"
Get-VM "v$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v1name )

$item.vvirt = 1
$over++
} # if
if ( $over -lt $threshold -and $item.ivirt -eq 9 ) {
write-host " moving i$($item.name) to  $($item.v1name)"
Get-VM "i$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v1name )

$item.ivirt = 1
$over++
} # if
if ( $over -lt $threshold -and $item.pvirt -eq 9 ) {
write-host " moving p$($item.name) to  $($item.v1name)"
Get-VM "p$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v1name )

$item.pvirt = 1
$over++
} # if
if ( $over -lt $threshold -and $item.xvirt -eq 9 ) {
write-host " moving x$($item.name) to  $($item.v1name)"
Get-VM "x$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v1name )

$item.xvirt = 1
$over++
} # if
if ( $over -lt $threshold -and $item.fvirt -eq 9 ) {
write-host " moving f$($item.name) to  $($item.v1name)"
Get-VM "f$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v1name )
$item.vvirt = 1
$over++
} # if
if ( $over -lt $threshold -and $item.qvirt -eq 9 ) {
write-host " moving q$($item.name) to  $($item.v1name)"
Get-VM "q$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v1name )

$item.qvirt = 1
$over++
} # if

} #if


if ($item.GoodV2 -eq 1) {
#v2 set
$over = $item.v2vcount

if ( $over -lt $threshold -and $item.vvirt -eq 9 ) {
# this is where the magic command happens to vmotion the server over!
Get-VM "v$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v2name )
$item.vvirt = 2
$over++
} # if
if ( $over -lt $threshold -and $item.ivirt -eq 9 ) {

Get-VM "i$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v2name )
$item.ivirt = 2
$over++
} # if
if ( $over -lt $threshold -and $item.pvirt -eq 9 ) {

Get-VM "p$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v2name )
$item.pvirt = 2
$over++
} # if
if ( $over -lt $threshold -and $item.xvirt -eq 9 ) {

Get-VM "x$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v2name )
$item.xvirt = 2
$over++
} # if
if ( $over -lt $threshold -and $item.fvirt -eq 9 ) {

Get-VM "f$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v2name )
$item.vvirt = 2
$over++
} # if
if ( $over -lt $threshold -and $item.qvirt -eq 9 ) {

Get-VM "q$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v2name )
$item.qvirt = 2
$over++
} # if

} #if





if ($item.GoodV3 -eq 1) {
#v3 set
$over = $item.v3vcount

if ( $over -lt $threshold -and $item.vvirt -eq 9 ) {
# this is where the magic command happens to vmotion the server over!
Get-VM "v$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v3name )
$item.vvirt = 3
$over++
} # if
if ( $over -lt $threshold -and $item.ivirt -eq 9 ) {

Get-VM "i$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v3name )
$item.ivirt = 3
$over++
} # if
if ( $over -lt $threshold -and $item.pvirt -eq 9 ) {

Get-VM "p$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v3name )
$item.pvirt = 3
$over++
} # if
if ( $over -lt $threshold -and $item.xvirt -eq 9 ) {

Get-VM "x$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v3name )
$item.xvirt = 3
$over++
} # if
if ( $over -lt $threshold -and $item.fvirt -eq 9 ) {

Get-VM "f$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v3name )
$item.vvirt = 3
$over++
} # if
if ( $over -lt $threshold -and $item.qvirt -eq 9 ) {

Get-VM "q$($item.name)" | Move-VM -Destination ( Get-VMHost $item.v3name )
$item.qvirt = 3
$over++
} # if

} #if


$i++
} # foreach


} # if #test block