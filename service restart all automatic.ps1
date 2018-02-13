param (

[Parameter(Mandatory=$True,ValueFromPipeline=$true)]
$server

)

#gets inital list of services
$servicelist = Get-Service -ComputerName $server


Write-Host "All services on $($server) that are Automatic but not running:"

foreach ($service in $servicelist) {
if ($service.StartType -eq "Automatic" -and $service.Status -ne "Running" ) {

#$betterlist += $service
Write-host $service.name $service.starttype $service.status



} #if
} #foreach


Write-Host "Starting listed services"

foreach ($service in $servicelist) {
if ($service.StartType -eq "Automatic" -and $service.Status -ne "Running" ) {

$service | Start-Service



} #if
} #foreach


write-Host "All services on $($server) that are Automatic but not running after start, and their new status:"

foreach ($service in $servicelist) {
if ($service.StartType -eq "Automatic" -and $service.Status -ne "Running" ) {

#$betterlist += $service
Write-host $service.name $service.starttype $service.status
} #if
} #foreach