# set directory and create directory structure.

$tdir = Test-Path C:\autodoc -PathType Container

If ($tdir -eq $false){
        
        $mdir = New-Item -Path c:\autodoc -Force -ItemType Diretory
}
$tdirf = Test-Path c:\autodoc\final -PathType Container

If ($tdirf -eq $false){
    $mdirf = New-Item -Path c:\autodoc\final -ItemType Directory -Force
}

# change working directory of script

Set-Location -Path C:\autodoc -PassThru

# clone the git
$alias = Get-Alias -Name git

If($alias -eq $false){
    
    New-Alias -Name git -Value "$Env:ProgramFiles\Git\bin\git.exe"
}

try{

    git clone https://github.com/mve83/teamsdoc 

  }catch{

    }

# set pandoc alias

$alias = Get-Alias -Name pandoc

If($alias -eq $false){
    
    New-Alias -Name pandoc -Value "$Env:ProgramFiles\pandoc\pandoc.exe"
}

# get parameters file and load variables for script

$file = Get-Content  .\teamsdoc\params.json | ConvertFrom-Json

ForEach($var in $file.variables){
        
        Set-Variable -Name $var.param -Value $null
}

## Set variable values
$customer = Read-Host "Please enter customer name"
$supplier = "FITTS"
$date = Get-Date -Format d-mm-y
$plan = Read-Host "Enter O365 Plan e.g. E5"
$version = Read-Host "Enter Document Version"
$residency = Read-Host "Enter Tenant Location"
$audioconferencinglicenses = Read-Host "Enter how many audio conferencing licenses required"
$azureadp1licenses = Read-Host "Enter how many Azure P1 licenses needed"
$meetingroomlicenses = Read-Host "Enter how many meeting room licenses needed"
$commonarealicenses = Read-Host "Enter how many common area phone licenses required"
$domcallingplanlicenses = Read-Host "Enter how many Domestic Calling Plan licenses required"
$intcallingplanlicenses = Read-Host "Enter how many International Calling Plan licenses required"
$phonesystemlicenses = Read-Host "Enter how many phone system licenses required"
$enterpriseuserlicenses = Read-Host "Enter how many E plans are required"
$communicationcredits = Read-Host "How much will be loaded into Communication Credits?"
$virtualuserlicenses = Read-Host "How many virtual user licenses required"
$documentreference = "fittsref.docx"
## Create YAML file

$yaml = ".\teamsdoc\docmeta.yaml"

$filetest = Test-Path .\teamsdoc\docmeta.yaml

if ($filetest -eq $true){
    
    Remove-Item -Path .\teamsdoc\docmeta.yaml -Force -Confirm:$false
}

New-Item -Path $yaml -ItemType File -Force

$content = "customer: $($customer)
supplier: $($supplier)
date:  '$($date)'
plan: $($plan)
version: '$($version)'
residency: $($residency)
audioconferencinglicenses: '$($audioconferencinglicenses)'
azureadp1licenses: '$($azureadp1licenses)'
meetingroomlicenses: '$($meetingroomlicenses)'
commonarealicenses: '$($commonarealicenses)'
domcallingplanlicenses: '$($domcallingplanlicenses)'
intcallingplanlicenses: '$($intcallingplanlicenses)'
phonesystemlicenses: '$($phonesystemlicenses)'
enterpriseuserlicenses: '$($enterpriseuserlicenses)'
communicationcredits: '$($communicationcredits)'
requiredvirtualuserlicenses: '$($virtualuserlicenses)'
virtualuserlicenses: '10'"

Set-Content -Path $yaml -Value $content

## Create document template

Function Generate-DocTemplate{
        
        param (
            [string]$Template = 'cloudcollab'
        )
       
       $doc = $file.$Template

       foreach($element in $doc){
            
            $input = "$($input) $($element.file)"
       }

       $mdfiles = $input.Substring(1)

       Set-Location -Path C:\autodoc\teamsdoc

       $test = Test-Path -Path command.cmd

       If ($test -eq $true){

        Remove-Item -Path command.cmd -Force -confirm:$false

       }

       New-item -ItemType file -name command.cmd

       set-content -path command.cmd -value "cd c:\autodoc\teamsdoc
       pandoc.exe $($mdfiles) --filter pandoc-mustache --toc --standalone --reference-doc $($documentreference) --o LLD.docx"

       .\command.cmd

       # move lld to final folder

       Move-Item -Path LLD.docx -Destination C:\autodoc\final

       Set-Location -Path 'C:\autodoc'

       Rename-Item -Path .\final\LLD.docx -NewName "$($customer) Teams Low Level Design.docx"

       ## clean up directory

       Remove-Item -Path .\teamsdoc -Recurse -Force -Confirm:$false


       Write-Host "Finished creating document, check c:\autodoc\final for the document. You may now press Q to quit..." -ForegroundColor Green
       
}
## set option menu

function Show-Menu
{
     param (
           [string]$Title = 'Please choose which type of document to create'
     )
     cls
     Write-Host "================ $Title ================"
     
     Write-Host "1: Press '1' Collab Only."
     Write-Host "2: Press '2' Collab and Meetings."
     Write-Host "3: Press '3' Cloud PSTN Calling."
     Write-Host "4: Press '4' Direct Routing."
     Write-Host "5: Press '5' Hybrid PSTN"
     Write-Host "Q: Press 'Q' to quit."
}

do
{
     Show-Menu
     $input = Read-Host "Please make a selection"
     switch ($input)
     {
           '1' {
                cls
                'You chose Collab Only'
                Generate-DocTemplate -Template "cloudcollab"
           } '2' {
                cls
                'You chose Collab and Meetings'
                Generate-DocTemplate -Template "cloudcollabmeetings"
           } '3' {
                cls
                'You chose Cloud PSTN Calling'
                Generate-DocTemplate -Template "cloudcalling"
           } '4' {
                cls
                'You chose Direct Routing - To be added'
           } '5' {
                cls
                'You chose Hybrid Voice - To be added'
           } 'q' {
                return
           }
     }
    pause

}
until ($input -eq 'q')

