<# 
Date created 5/23/2019
Created by Audrey Sin Fai Lam

Purpose: Takes in the VSE_Current_DAT_Adoption.csv file. Adds or updates the status of the servers with DAT version out-of-sync
Requires:Input file is in csv format
#>


<# *******************************************************************************************************************
Note:
Requires input file is in csv format

Input file:
VSE_Current_DAT_Adoption.csv

By-product files:
VSE_Current_DAT_Adoption copy (date).csv (copy of the input saved in the directory)
VSE_Current_DAT_Adoption update (date).csv (copy of the input without the first 4 lines for processing purposes)

Output files:
McAfee cumulative incident log.csv (File which accumulates the daily out-of-sync records)
McAfee incident out-of-sync (date) .csv (File generated with the date's out-of-sync records if any)

********************************************************************************************************************** #>

$folder = "Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check"

# ******
#test:
# $folder = "Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\Coop Handover- Audrey McAfee Scripts\Update\test output"
# ****

$file_collection = get-childitem -path $folder

Echo 'Taking in inputs'

# Input path:
$inputpath = "Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\VSE_Current_DAT_Adoption.csv"

# By-product paths
$historypath = ("Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\VSE_Current_DAT_Adoption History\VSE_Current_DAT_Adoption #" + ((get-date).ToString('MM-dd-yy')) +  ".csv")
$updatepath =  ("Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\VSE_Current_DAT_Adoption History\Archive\VSE_Current_DAT_Adoption update " + ((get-date).ToString('MM-dd-yy')) +  ".csv")

 
# Output paths
# 1) Generated daily if there is any incident to be raised. If so, the file is sent by email to raise the incident.
$dailyincidentpath = "Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\McAfee reports\Daily\Incidents\McAfee incident out-of-sync " + ((get-date).ToString('MM dd yy')) +  ".csv"
# 2) Systems out-of-sync is added to this file daily
$cumulativepath = "Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\McAfee cumulative incident log.csv"
# 3) This is a copy of file 2) to be kept for record
$cumulative_history = "Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\McAfee reports\Daily\McAfee cumulative incident log " + ((get-date).ToString('MM-dd-yy')) +  ".csv"


#test
<#
# Input path:
$inputpath = "Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\Coop Handover- Audrey McAfee Scripts\Update\test input\VSE_Current_DAT_Adoption.csv"

# By-product paths
$historypath = ("Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\Coop Handover- Audrey McAfee Scripts\Update\test input\VSE_Current_DAT_Adoption #" + ((get-date).ToString('MM-dd-yy')) +  ".csv")
$updatepath =  ("Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\Coop Handover- Audrey McAfee Scripts\Update\test input\VSE_Current_DAT_Adoption update " + ((get-date).ToString('MM-dd-yy')) +  ".csv")
 

# Output paths
# 1) Generated daily if there is any incident to be raised. If so, the file is sent by email to raise the incident.
$dailyincidentpath = "Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\Coop Handover- Audrey McAfee Scripts\Update\test output\McAfee incident out-of-sync " + ((get-date).ToString('MM dd yy')) +  ".csv"
# 2) Systems out-of-sync is added to this file daily
$cumulativepath = "Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\Coop Handover- Audrey McAfee Scripts\Update\test output\McAfee cumulative incident log.csv"
# 3) This is a copy of file 2) to be kept for record
$cumulative_history = "Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\Coop Handover- Audrey McAfee Scripts\Update\test output\McAfee cumulative incident log " + ((get-date).ToString('MM-dd-yy')) +  ".csv"
#>

# *************************************************************************************************************************************

# This variable can be maually modified depending on the number of lines that can be skipped at the top of VSE_Current_DAT_Adoption.csv file

$skip_lines = 4


# ***********************************************************************************************************************************

# Saves a copy to History folder
get-content $inputpath | set-content $historypath

# Creates a copy of the file withoput the top 4 lines
get-content $inputpath | select -skip $skip_lines | set-content $updatepath

# Imports the input file
$daily_DAT_adoption = import-csv -path $updatepath

# Variable
$found_output = $false

# Checks if the output log file is existent
if ($file_collection.name -contains 'McAfee cumulative incident log.csv') {
    $found_output = $true
}


# Creates the output log file if non-existent
if ($found_output -eq $false) {
Add-content -path $cumulativepath -value 'Name,Last Communication,Incident Creation Date,Date of Most Recent Check,Date Remediated,Days out-of-sync,Remediated,DAT Version (VirusScan Enterprise),Required DAT Version,Operating System,Product Version (VirusScan Enterprise),Engine Version (VirusScan Enterprise)'
}

# ***********************************************************************************************************************************


# Variables
$days_out_of_sync_allowed = 2
$alert = $false
$unresolved_case = $false
$new_incident = @{}
$version_diff = $null
$version_diff_allowed = 1
$sys_name = $null
$DAT_version = $null
$required_DAT_ver = $null
$last_check = (get-date).toString('M/d/yy')
$days_out_of_sync = $null
$no_incident_or_remediation = 'N/A'
$resolved = 'No'
$p_version = $null
$Eng_version = $null
$last_communication = $null
$OS = $null


# *****************************************************************************************************************************************

# Mutates the global variable $required_DAT_ver to be the DAT version of OMS511. 
# If the OMS511 server is not in the list, the script stops and an email is sent to IT OPS team.

foreach ($sys in $daily_DAT_adoption) {
    if ($sys.'System name' -contains 'OMS511') {
        $required_DAT_ver = $sys.'DAT Version (VirusScan Enterprise)'
        Echo "It's here"
    }
}


if ($required_DAT_ver -eq $null) {
    $msg = 
"Good morning,
    
Please note that an error was encountered while running the McAfee script. Please check that the server OMS511 is in the VSE_Current_DAT_Adoption.csv and check that the number of rows at the top of the VSE_Current_DAT_Adoption.csv file.

Regards,

IT Operations and Vendor Management"

    Send-MailMessage -From 'itoperations&vendormanagement@Omers.com' -To 'itoperations&vendormanagement@Omers.com' -Subject 'OMS511 missing' -Body $msg -SmtpServer "mail.omers.com"
    break
    # exits the script
}


# ********************************************************************************************************************************************

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
                        $system.'Date of Most Recent Check'= $last_check
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
                

                $new_incident = New-Object PsObject -Property @{'Name' = $sys_name
                'Last Communication' = $last_communication
                'Incident Creation Date' = $no_incident_or_remediation
                'Date of Most Recent Check' = $last_check
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
                        $item.'Date of Most Recent Check' = $last_check
                        $item.'DAT Version (VirusScan Enterprise)' = $DAT_version
                        $item.'Required DAT Version' = $required_DAT_ver
                        $item.'Operating System' = $OS
                        $item.'Product Version (VirusScan Enterprise)' = $p_version
                        $item.'Engine Version (VirusScan Enterprise)' = $Eng_version
                        $item.'Remediated' = 'Yes'
                        $item.'Date Remediated' = (get-date).toString('M/d/yy')
                        [pscustomobject] $cumulative_file | export-csv -path $cumulativepath -NoTypeInformation
                        
                    }
             }
         }
    $unresolved_case = $false
    $new_incident = @{}

    $sys_name = $null
    $DAT_version = $null
    
    $days_out_of_sync = $null
    
    $version_diff = $null
    $p_version = $null
    $Eng_version = $null
    $last_communication = $null
    $OS = $null
    }
}



# **************************************************************************************************************************************************

# Newly added

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

                            
$cumulative_file = import-csv -path $cumulativepath

foreach ($object in $cumulative_file) {
    if (($object.'Days out-of-sync' -gt $days_out_of_sync_allowed) -and
        (($object.'Remediated' -contains 'N/A') -or 
         ($object.'Remediated' -contains 'No'))) {
        $object.'Incident Creation Date' = $last_check
        if ($alert -eq $false) {
            Add-content -path $dailyincidentpath -value 'Name,Last Communication,DAT Version (VirusScan Enterprise),Required DAT Version,Operating System,Product Version (VirusScan Enterprise),Engine Version (VirusScan Enterprise)'
            $alert = $true
        }
        $object.'Remediated' = 'No'

        $new_incident = 
        @{'Name' = $object.Name
        'Last Communication' = $object.'Last Communication'
        'DAT Version (VirusScan Enterprise)' = $object.'DAT Version (VirusScan Enterprise)'
        'Required DAT Version' = $object.'Required DAT Version'
        'Operating System' = $object.'Operating System'
        'Product Version (VirusScan Enterprise)' = $object.'Product Version (VirusScan Enterprise)'
        'Engine Version (VirusScan Enterprise)' = $object.'Engine Version (VirusScan Enterprise)'
        }
        [pscustomobject] $new_incident | export-csv -path $dailyincidentpath -Append -NoTypeInformation
        [pscustomobject] $cumulative_file | export-csv -path $cumulativepath -NoTypeInformation
    }
}

# ************************************************************************************************************************************************************

# Makes a daily back-up copy of the cumulative log file
get-content $cumulativepath | set-content $cumulative_history

# ************************************************************************************************************************************************************


Write-host "Adding inventory info"
import-module -name "Z:\Server Inventory\McAfee Report\McAfee Non Compliant Check\Coop Handover- Audrey McAfee Scripts\Add inventory info_McAfee.psm1"
$list1 = @("short_description","u_environment","hardware_status","support_group","department","u_business_entity")


inv_info_appender2 $cumulative_history $list1

if ($alert) {
    write-host "Adding inventory info to incident file"
    inv_info_appender2 $dailyincidentpath $list1

}

# Email text       
$incident_text = "Hi,

Please find attached the servers out-of-sync by more than " + $days_out_of_sync_allowed.toString() + " days.
Kindly investigate into the servers that are in deployed state.


Regards,

IT Operations and Vendor Management"

#Sends email to raise incident to server team if there are servers out-of-sync

if ($alert) {
Send-MailMessage -From 'itoperations&vendormanagement@Omers.com' -To 'itoperations&vendormanagement@Omers.com' -Cc 'itoperations&vendormanagement@Omers.com' -Subject 'McAfee signature out-of-sync' -Body $incident_text -Attachment $dailyincidentpath -SmtpServer "mail.omers.com"
Write-host 'Incident alert sent'
Echo $alert
$alert = $false
} #>

# Email text
$daily_msg = "Hi,

Please find attached the McAfee cumulative report as of today.

Regards,

IT Operations and Vendor Management"

# Send email daily
Send-MailMessage -From 'itoperations&vendormanagement@Omers.com' -To 'itoperations&vendormanagement@Omers.com' -Subject 'McAfee signature cumulative report' -Body $daily_msg -Attachments $cumulative_history -SmtpServer "mail.omers.com"
Write-host 'Log file sent'

# ***************************************************************************************************

Echo "Complete!"