# Date 5/23/2019


<# *******************************************************************************************************************
Note:
Requires input file is in csv format

Input file:
VSE_Current_DAT_Adoption (date).csv

By-product files:
VSE_Current_DAT_Adoption update (date).csv (copy of the input without the first 4 lines for processing purposes)

Output files:
McAfee cumulative incident log.csv (File which accumulates the daily out-of-sync records)

Assumptions:
1) The input files are not modified at any point after they have been placed in the directory
   (i.e. the 'Last date modified' matches the date of creation of the file)
2) No file is missing from the folder
********************************************************************************************************************** #>

# Output paths
# Systems out-of-sync is added to this file
$output_file = "McAfee batch cumulative incident log " + (get-date).toString('M-d-yy') + ".csv"
# The following line is commented out for testing purposes
$output_folder = "Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\McAfee reports\Batch"

# Test
#$output_folder = "Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\Coop Handover- Audrey McAfee Scripts\Update\test output\Batch"

$cumulativepath = $output_folder + '\' + $output_file

# Variable
$found_output = $false

$file_collection = get-childitem -path $output_folder

# Checks if the output log file is existent
if ($file_collection.name -contains $output_file) {
    $found_output = $true
}

# ***********************************************************************************************************************

# Variables
$days_out_of_sync_allowed = 2
$unresolved_case = $false
$new_incident = $null
$version_diff = $null
$version_diff_allowed = 1
$sys_name = $null
$DAT_version = $null
$required_DAT_ver = $null
$days_out_of_sync = $null
$no_incident_or_remediation = 'N/A'
$resolved = 'No'
$p_version = $null
$Eng_version = $null
$last_communication = $null
$OS = $null
$date = $null
$OMS511_missing = $null

# ************************************************************************************************************************

# Creates the output log file if non-existent
if ($found_output -eq $false) {
Add-content -path $cumulativepath -value 'Name,Last Communication,Incident Creation Date,Date of Most Recent Check,Date Remediated,Days out-of-sync,Remediated,DAT Version (VirusScan Enterprise),Required DAT Version,Operating System,Product Version (VirusScan Enterprise),Engine Version (VirusScan Enterprise)'
}

# *********************************************************************************************************************

$directory = "Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\VSE_Current_DAT_Adoption History"

$DAT_files = get-childitem -path ($directory + '\*.csv')  | sort LastWriteTime



foreach ($file in $DAT_files) {

    # Input path:
    $inputpath = $directory + '\' + $file.Name

    # By-product paths
    $updatepath =  ("Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\VSE_Current_DAT_Adoption History Update\" + ($file.Name).trim(".csv") +  "update.csv")


    # ***********************************************************************************************************************************

    Echo ('Reading ' + $file.Name)

    # Creates a copy of the file withoput the top 4 lines
    get-content $inputpath | select -skip 4 | set-content $updatepath

    # Imports the input file
    $daily_DAT_adoption = import-csv -path $updatepath
    # Test $daily_DAT_adoption = import-csv -path $inputpath

    # ****************************************************************************************************************************************

    # Get file date

    $separator = "#" 
    $parts = ($file.Name).split($separator)
    $date = ($parts[1]).trim(".csv")
   

    # *****************************************************************************************************************************************

    # Mutates the global variable $required_DAT_ver to be the DAT version of OMS511. 
    # If the OMS511 server is not in the list, the script stops and an email is sent to IT OPS team.

    foreach ($sys in $daily_DAT_adoption) {
        if ($sys.'System name' -contains 'OMS511') {
            $required_DAT_ver = $sys.'DAT Version (VirusScan Enterprise)'
        }
    }


    if ($required_DAT_ver -eq $null) {
        $OMS511_missing += $date + " "
        Write-Host ("OMS511 missing on " + $date)
        continue
        # goes to the next file
    }


    # *********************************************************************************************************************************************


    # Check each server in the daily DAT file
    foreach ($server in $daily_DAT_adoption) {
        # skips rows with no server
        if ($server."Last Communication" -contains "Last Communication") {
            continue

        } else {

            $cumulative_file = import-csv -path $cumulativepath

            $sys_name = $server.'System name'
            $DAT_version = $server.'DAT Version (VirusScan Enterprise)'
            $p_version = $server.'Product Version (VirusScan Enterprise)'
            $Eng_version = $server.'Engine Version (VirusScan Enterprise)'
            $last_communication = $server.'Last Communication'
            $OS = $server.'Operating System'

            # ***********************************************************************************************

            # If the DAT version is out-of-sync (DAT version is considered out-sync if it is more than one version behind the of OMS511)
            $version_diff = ([int]$required_DAT_ver) - ([int]($server."DAT Version (VirusScan Enterprise)"))
            if ($version_diff -gt $version_diff_allowed) {

                foreach ($ci in $cumulative_file) {
                    if (($ci.Name -contains $server.'System name') -and
                        (($ci.Remediated -contains 'No') -or
                         ($ci.Remediated -contains 'N/A'))) {
                         $unresolved_case = $true
                    }
                }
            
                # If there is an open issue for this server, update the latest open entry

                if ($unresolved_case -eq $true) {
                    foreach ($system in $cumulative_file) {
                        if (($system.Name -eq $server.'System name') -and
                            (($system.Remediated -contains 'No') -or
                             ($system.Remediated -contains 'N/A'))) {
                            $system.'Date of Most Recent Check'= $date
                            $system.'Days out-of-sync' = [int]($system.'Days out-of-sync') + 1
                            $system.'Last Communication' = $last_communication
                            $system.'DAT Version (VirusScan Enterprise)' = $DAT_version
                            $system.'Required DAT Version' = $required_DAT_ver
                            $system.'Operating System' = $OS
                            $system.'Product Version (VirusScan Enterprise)' = $p_version
                            $system.'Engine Version (VirusScan Enterprise)' = $Eng_version

                            [pscustomobject] $cumulative_file | export-csv -path $cumulativepath -NoTypeInformation
                            $unresolved_case = $false
                            break
                        }
                    }

                } else {

                    # This is the first time an issue is detected for this server or there is no unresolved issue for this server
                

                    [pscustomobject] $new_incident = New-Object PsObject -Property @{'Name' = $sys_name
                    'Last Communication' = $last_communication
                    'Incident Creation Date' = $no_incident_or_remediation
                    'Date of Most Recent Check' = $date
                    'Date Remediated' = $no_incident_or_remediation
                    'Days out-of-sync' = 1
                    'Remediated' = $no_incident_or_remediation
                    'DAT Version (VirusScan Enterprise)' = $DAT_version
                    'Required DAT Version' = $required_DAT_ver
                    'Operating System' = $OS
                    'Product Version (VirusScan Enterprise)' = $p_version
                    'Engine Version (VirusScan Enterprise)' = $Eng_version
                    }
                     [pscustomobject] $new_incident | export-csv -path $cumulativepath -Append -NoTypeInformation
                }

           
            
             } else {
                # The server is in-sync
                # If there is any open issue for this server, the record is updated
                foreach ($item in $cumulative_file) {
                        if (($item.Name -eq $server.'System name') -and
                            (($item.Remediated -eq 'N/A') -or
                             ($item.Remediated -eq 'No'))) {
                                               
                            $item.'Last Communication' = $last_communication
                            $item.'Date of Most Recent Check' = $date
                            $item.'DAT Version (VirusScan Enterprise)' = $DAT_version
                            $item.'Required DAT Version' = $required_DAT_ver
                            $item.'Operating System' = $OS
                            $item.'Product Version (VirusScan Enterprise)' = $p_version
                            $item.'Engine Version (VirusScan Enterprise)' = $Eng_version
                            $item.'Remediated' = 'Yes'
                            $item.'Date Remediated' = $date
                            [pscustomobject] $cumulative_file | export-csv -path $cumulativepath -NoTypeInformation
                        
                        }
                 }
             }
        $unresolved_case = $false
        $new_incident = $null
        $sys_name = $null
        $DAT_version = $null
        $days_out_of_sync = $null
        $p_version = $null
        $Eng_version = $null
        $last_communication = $null
        $OS = $null
        $version_diff = $null
        
        }
    }
                            
    $cumulative_file = import-csv -path $cumulativepath

    foreach ($object in $cumulative_file) {
        if (($object.'Days out-of-sync' -gt $days_out_of_sync_allowed) -and
            ($object.'Incident Creation Date' -contains 'N/A')) {
            $object.'Incident Creation Date' = $date
            $object.'Remediated' = 'No'
            [pscustomobject] $cumulative_file | export-csv -path $cumulativepath -NoTypeInformation
        }
    }
    $date = $null
    $required_DAT_ver = $null
    
}


function server_look_up ($server_name,$DAT_file) {
    foreach ($record in $DAT_file) {
        $server = $record.'System name'

        if ($server_name -contains $server) {
                return $true
                break
        }
    }
    return $false
}

$cumulative_file = import-csv -path $cumulativepath

foreach ($system_elem in  $cumulative_file) {
    

    if (($system_elem.Remediated -contains 'No') -or
         ($system_elem.Remediated -contains 'N/A')) {
         
        if ((server_look_up $system_elem.Name $daily_DAT_adoption) -eq $false) {
        write-host("Checking " + $system_elem.Name)
            $system_elem.'Remediated' = "Server removed"
            [pscustomobject] $cumulative_file | export-csv -path $cumulativepath -NoTypeInformation
        }
    }
}


Write-host "Adding inventory info"
import-module -name "Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\Coop Handover- Audrey McAfee Scripts\Update\Add inventory info_McAfee.psm1"
$list1 = @("short_description","u_environment","hardware_status","support_group","department","u_business_entity")


inv_info_appender2 $cumulativepath $list1

# ************************************************************************************************************************************************************

# Send an email if OMS511 is missing on any day

if ($OMS511_missing -ne $null) {
$msg = 
"Good morning,
    
Please note that the server 'OMS511' is missing from VSE_Current_DAT_Adoption.csv file on " + $OMS511_missing + "

Regards"

Send-MailMessage -From 'aulam@omers.com' -To 'aulam@omers.com' -Subject 'OMS511 missing' -Body $msg -SmtpServer "mail.omers.com"
}

# *************************************************************************************************************************************************************

Echo 'Complete'


