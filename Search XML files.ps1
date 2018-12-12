# Search through xml file for a string(s)
#
# By Frazer Grant date 28/02/2018 


# Version 1.6 - Added variable $heavyLoad, 

# useful note:
# using regex to search through files for variables, however not useful for this solution.
# Select-String -Pattern "\b(test|test2)W+(?:\w+\W+){1,10}?K1\b"

# If heavy load = 1 find the files on remote server and then copy them to local PC.  Set to 0 runs a command on remote server to find the files, zip them and then copy to local PC
$heavyLoad = 1


# 
$items = "Test1|Test2|Test3"

# set the date range reduce the amount of files search through.
$StartDate = "20181001"
$EndDate = "20181020"

$RemoteComputer = "prod.local"

$RemoteScrDir = "C:\RemoteXMLFiles"
$RemoteBackup = "H:\Backups\"




$LocalDir = "C:\temp\Test"
$LocalSrcDir = $LocalDir + "\" + $StartDate + "-" + $EndDate


# if folder does not exist then create it
if(!(Test-Path $LocalDir)){
    New-Item -ItemType directory -Path $LocalDir} 

# if folder does not exist then create it
if(!(Test-Path $LocalSrcDir)){
    New-Item -ItemType directory -Path $LocalSrcDir} 





if($heavyLoad -eq 1){ 
   
    
    Write-Output "Discovering XML Files"                                   
    (measure-command {$FileList = Get-ChildItem ("\\"+$RemoteComputer+"\"+$RemoteScrDir.replace(":","$")) -File -Filter *.xml | Where-Object {($_.CreationTime.ToString('yyyyMMdd') -ge $StartDate -and $_.CreationTime.ToString('yyyyMMdd')  -le $EndDate)}   }).TotalSeconds 
                      $i = 1                                     
                      $FileList                                      
                      write-host ""
                      (Measure-Command {foreach ($file in $FileList){
                        write-host "Copying " $i "out of" $FileList.count  " "  $File.FullName                                     
                        Copy-Item $File.Fullname $LocalSrcDir
                        $i++ }}).TotalSeconds 
                } 

 Elseif ($heavyLoad -eq 0) { 
   Invoke-Command –ComputerName $RemoteComputer –ScriptBlock { param($StartDate,$EndDate,$FileList ,$RemoteScrDir, $RemoteBackup)
                                                            Write-Output "Finding XML Files"
                                                           
                                                            (measure-command {$FileList = Get-ChildItem $RemoteScrDir -File -Filter *.xml | Where-Object {($_.CreationTime.ToString('yyyyMMdd') -ge $StartDate -and $_.CreationTime.ToString('yyyyMMdd')  -le $EndDate)} }).TotalSeconds 
                                                            $FileList
                                                            Write-Output "Creating Zip file containing XML files"
                                                            (measure-command {$FileList | compress-Archive  -CompressionLevel Fastest -DestinationPath $RemoteBackup"\ForSearching" -Force}).TotalSeconds
                                                            
                                                           } -ArgumentList $StartDate,$EndDate, $FileList ,$RemoteScrDir, $RemoteBackup



Write-Output "Copying Zip file to local PC"
Move-item ("\\"+$RemoteComputer+"\"+$RemoteBackup.replace(":","$")+"\ForSearching.zip") $LocalDir -Force



Write-Output "Extracting files"
Expand-Archive -LiteralPath $LocalDir"\ForSearching.zip" -DestinationPath $LocalSrcDir -Force

}



# create a folder called found, the results file will be placed here
$Found = $LocalSrcDir + '\' + 'Found'

# path to where the Amazon XML files are stored on Control Tower
$PCCTpath = "\\1.2.3.4\XML\"

$MacCTpath = "smb://1.2.3.4/XML/"

    
# create a list of all files in folder
$filesInDir = Get-ChildItem $LocalSrcDir -file -Filter *.xml

$i= 1

# if folder does not exist then create it
if(!(Test-Path $Found)){
    New-Item -ItemType directory -Path $Found} 

# create a blank object
$results = @()

# check each file in folder
(Measure-Command {foreach ($file in $filesInDir){
    
        write-host "reading " $i "out of" $filesInDir.count  " "  $file.FullName 
        # read each file into memory
        [xml]$xml = Get-Content $file.FullName
        # all useful info is stored under Node1/Node2/Node3
        $f = $xml.SelectNodes("Node1/Node2/Node3")
        
        # check each node in the XML object
        foreach ($items in $f){
            # Find parcel 
            if ($items.Node1.XMLNum -imatch $items){
          
             # once found parcel add details to an object
             $details = [Ordered]@{            
                Items           = $items.ID.XMLNum            
                Status           = $items.StatusInfo.Status                                 
                Date             = $file.lastWriteTime.ToString('yyyy-MM-dd')
                PCFilename       = "=HYPERLINK("+ [char]34 + $PCCTpath + $file.Name.ToString() + [char]34 + ","+ [char]34 + $file.Name.ToString() + [char]34 + ")"
               MacFilename       = "=HYPERLINK("+ [char]34 + $MacCTpath + $file.Name.ToString() + [char]34 + ","+ [char]34 + $file.Name.ToString() + [char]34 + ")"
               FileLengthInKB        = $file.Length / 1KB
             }                           
         # repeat adding found parcels to object    
        $results += New-Object PSObject -Property $details  
        }
        }
        $i++          
     }
}).TotalSeconds

$a = get-date
#$a.tostring("hh.mmdd-MM-yyyy-")

# export results object to csv file.
(Measure-Command{ $results | export-csv -Path ($Found + '\' + 'Test'+ ($a.tostring("hh.mm-dd-MM-yyyy-")) + 'foundXML.csv') -NoTypeInformation -Append   }).Totalseconds    
        


