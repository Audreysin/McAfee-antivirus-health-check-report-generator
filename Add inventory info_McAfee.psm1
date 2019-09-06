<#
Date: June 04, 2019
Created by Audrey Sin Fai Lam
Purpose: This script takes in a csv file (path:$client_file_path) with a column containing servers and appends inventory information ($list) corresponding to the servers.
         If a server is not in the inventory, "not in inventory" is entered in the in the fields
Requires: The input files are in csv format
          The column header names in $list must match the corresponding column header name in the inventory file (path: $inventory_path)
          One of the columns of the input file (path:$client_file_path) must contain server names

#>

# *************************************************************************************************************************************************

<#
Input paths:
$client_file_path
$inventory_path = "\\oms042\ProdDocs\ITReports\Enterprise Server Inventory Release.csv"

Output path:
$client_file_path

List of column headers to be appended:
$list

Header containing the server names in the $client_file_path:
$target_header

#>


# ************************************************************************************************************************************************
# Variables
$look_up_result = "This is a random value"
$inv_entry = "This is a random value"
$inventory_path = $null
$inventory = $null

# *************************************************************************************************************************************************

# Inventory path
$inventory_path = "\\oms042\ProdDocs\ITReports\Enterprise Server Inventory Release.csv"
$inventory = import-csv -path $inventory_path
Write-host $inventory_path


# ************************************************************************************************************************************************

<#
Purpose: Looks up in the inventory if $server_name is in the inventory.
         If so, it returns the whole entry for the server.
         Otherwise, it returns "Not in ineventory"
Requires: The column heading containing the server names in the inventory file is "name"
Effects: Mutate global variable $look_up_result
#>

function inventory_look_up ($server_name,$inventory_file) {
    $look_up_result = 'Not in inventory'
    foreach ($record in $inventory_file) {
        $server = $record.name

        if (($server_name -contains $server) -or
            ($server_name -like ($server + ".*")) -or
            ($server -like ($server_name + ".*")) -or
            ($server_name -like ($server + "-*")) -or
            ($server -like ($server_name + "-*"))) {
             
                $look_up_result = $record
                break
        }
    }
    return $look_up_result
}

# **********************************************************************************

<#
Purpose: Appends the required columns to the csv file
#>

function inv_info_appender2($file_path,$list){
	$file = import-csv -path $file_path
	foreach ($item in $list) {
		Write-host ("Creating column " + $item)
		[pscustomobject]  $file | 
        Select-Object *,$item | 
        Export-Csv $file_path -NoTypeInformation
        $file = import-csv -path $file_path  #May be deletable
	}

	$file = import-csv -path $file_path
	foreach ($ci in $file) {
		Write-host ("working on " + $ci."Name")
		$inv_entry = inventory_look_up ($ci."Name") $inventory
		if ($inv_entry -eq "Not in inventory") {
			foreach ($item in $list) {
				$ci.$item = $inv_entry
			}

        } else {
        	foreach($item in $list) {
        		$ci.$item = ($inv_entry.$item).ToString()
        	}
            $look_up_result = $null
            [pscustomobject]  $file | export-csv $file_path -NoTypeInformation #Maybe
        }
        [pscustomobject]  $file | export-csv $file_path -NoTypeInformation
	}
}

Export-ModuleMember -Function inventory_look_up
Export-ModuleMember -Function inv_info_appender2

Export-ModuleMember -Variable $inventory_path
Export-ModuleMember -Variable $inventory
Export-ModuleMember -Variable $look_up_result
Export-ModuleMember -Variable $inv_entry


