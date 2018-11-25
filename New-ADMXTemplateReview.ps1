<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.153
	 Created on:   	18/09/2018 9:48 AM
	 Created by:   	Jeffrey Wadham
	 Organization: 	Queensland Health
	 Filename:     	New-ADMXTemplateReview.ps1
	===========================================================================
	.DESCRIPTION
		Three functions used to review ADMX files before they are imported into the Central Store.
		Export-ADMXInfo will translate ADMX files into a readable CSV
		Compare-ADMXCSV can then be used to find any settings present in the OLD ADMX that are not in the NEW ADMX (Deprecated)
		Search-GPObjects can then search all GPOs in the domain for a string, using the export from Compare-ADMXCSV you can find a Display Name to search for
#>

<#

	.DESCRIPTION
		Translates ADMX files using the ADML files into a readable CSV file
	
	.PARAMETER PolicyDir
		The directory containing your ADMX and ADML files (PolicyDefinitions)
	
	.PARAMETER Language
		The language used to translate the files e.g. en-us, en-es
	
	.PARAMETER OutputFileName
		The path and filename for the CSV export
	
	.EXAMPLE
				PS C:\> Export-ADMXInfo -PolicyDir "$($env:windir)\policyDefinitions" -Language en-US -OutputFileName "C:\Temp\ADMX.csv
	
	.NOTES
		As ADMX files can be initially installed into a different directory other than the local store this function can be used to compare new and old ADMX sets
#>
function Export-ADMXInfo
{
	param
	(
		[string]$PolicyDir,
		[string]$Language,
		[string]$OutputFileName
	)
	
	$table = New-Object System.Data.DataTable
	
	[void]$table.Columns.Add("ADMX")
	[void]$table.Columns.Add("ParentCategory")
	[void]$table.Columns.Add("Name")
	[void]$table.Columns.Add("DisplayName")
	[void]$table.Columns.Add("Class")
	[void]$table.Columns.Add("ExplainText")
	[void]$table.Columns.Add("SupportedOn")
	[void]$table.Columns.Add("Key")
	[void]$table.Columns.Add("ValueName")
	
	$admxFiles = Get-ChildItem $policyDir -filter *.admx
	$count = 1
	
	ForEach ($file in $admxFiles)
	{
		Write-Progress -id 1 -Activity "Processing files" -CurrentOperation $file -PercentComplete $($count/$($admxfiles.count) * 100) -Status "$count of $($admxFiles.count)"; $count++
		[xml]$data = Get-Content "$policyDir\$($file.Name)"
		[xml]$lang = Get-Content "$policyDir\$language\$($file.Name.Replace(".admx", ".adml"))"
		
		$policyText = $lang.policyDefinitionResources.resources.stringTable.ChildNodes
		
		$data.PolicyDefinitions.policies.ChildNodes | ForEach-Object {
			
			$policy = $_
			
			if ($policy -ne $null)
			{
				if ($policy.Name -ne "#comment")
				{
					"Processing policy $($policy.Name)"
					$displayName = ($policyText | Where-Object { $_.id -eq $policy.displayName.Substring(9).TrimEnd(')') }).'#text'
					$explainText = ($policyText | Where-Object { $_.id -eq $policy.explainText.Substring(9).TrimEnd(')') }).'#text'
					
					if ($policy.SupportedOn.ref.Contains(":"))
					{
						$source = $policy.SupportedOn.ref.Split(":")[0]
						$valueName = $policy.SupportedOn.ref.Split(":")[1]
						[xml]$adml = Get-Content "$policyDir\$language\$source.adml"
						$resourceText = $adml.policyDefinitionResources.resources.stringTable.ChildNodes
						$supportedOn = ($resourceText | Where-Object { $_.id -eq $valueName }).'#text'
						
					}
					else
					{
						$supportedOnID = ($data.policyDefinitions.supportedOn.definitions.ChildNodes | Where-Object { $_.Name -eq $policy.supportedOn.ref }).DisplayName
						$supportedOn = ($policyText | Where-Object { $_.id -eq $supportedOnID.Substring(9).TrimEnd(')') }).'#text'
					}
					
					if ($policy.parentCategory.ref.Contains(":"))
					{
						$source = $policy.SupportedOn.ref.Split(":")[0]
						$valueName = $policy.SupportedOn.ref.Split(":")[1]
						[xml]$adml = Get-Content "$policyDir\$language\$source.adml"
						$resourceText = $adml.policyDefinitionResources.resources.stringTable.ChildNodes
						$parentCategory = ($resourceText | Where-Object { $_.id -eq $valueName }).'#text'
						
					}
					else
					{
						$parentCategoryID = ($data.policyDefinitions.categories.ChildNodes | Where-Object { $_.Name -eq $policy.parentCategory.ref }).DisplayName
						$parentCategory = ($policyText | Where-Object { $_.id -eq $parentCategoryID.Substring(9).TrimEnd(')') }).'#text'
					}
					
					[void]$table.Rows.Add(
						$file.Name,
						$parentCategory,
						$policy.Name,
						$displayName,
						$policy.class,
						$explainText,
						$supportedOn,
						$policy.key,
						$policy.valueName)
				}
			}
			
		}
	}
	
	$table | Export-Csv $outputfilename -NoTypeInformation
}

<#
	.SYNOPSIS
		Compares two CSV exports of ADMX files
	
	.PARAMETER OldADMX
		Path to the CSV export of your old/current ADMX.
	
	.PARAMETER NewADMX
		Path to the CSV export of your new ADMX.
	
	.PARAMETER ExportPath
		The folder path you want to export the results to as a CSV.
	
	.PARAMETER Column
		The column you wish to use as your unique identifier.
	
	.EXAMPLE
		PS C:\> Compare-ADMXCSVs -Column 'Name' -OldADMX 'C:\Temp\ADMX\OldADMX.csv' -NewADMX 'C:\Temp\ADMX\NewADMX.csv' -ExportPath 'C:\Temp\ADMX\Deprecated.csv'
	
	.NOTES
		Based on the information exported by Export-ADMXInfo the best column to use as a unique identifier is "Name".
#>
function Compare-ADMXCSV
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$OldADMX,
		[Parameter(Mandatory = $true)]
		[string]$NewADMX,
		[Parameter(Mandatory = $true)]
		[string]$ExportPath,
		[Parameter(Mandatory = $true)]
		[string]$Column
	)
	
	$OldADMXCSV = Import-Csv -Path $OldADMX
	$NewADMXCSV = Import-Csv -Path $NewADMX
	$DeprecatedSettings = @()
	$count = 1
	## Begin checking each row in the Old ADMX CSV for deprecated settings
	foreach ($OldADMXCSVRow in $OldADMXCSV)
	{
		Write-Progress -id 1 -Activity "Processing table" -CurrentOperation $($OldADMXCSVRow.DisplayName) -PercentComplete $($count/$($OldADMXCSV.count) * 100) -Status "$count of $($OldADMXCSV.count)"; $count++
		## Find the row match in the New ADMX CSV from the displayname specified 
		$NewADMXCSVRow = $NewADMXCSV | where { $_.$Column -eq $OldADMXCSVRow.$Column }
		## If matches were found 
		if ($NewADMXCSVRow)
		{
			Write-Host "Matches found for $($OldADMXCSVRow.DisplayName) in NewADMX CSV"
			
		}
		else
		{
			Write-Host "No match found for $($OldADMXCSVRow.DisplayName) in NewADMX CSV"
			$DeprecatedSettings += $OldADMXCSVRow
			
		}
	}
	$DeprecatedSettings | Export-Csv "$ExportPath" -NoTypeInformation
}

<#
	.SYNOPSIS
		Searches all GPOs in a domain for a string.
	
	.DESCRIPTION
		Gets a XML report of all GPOs in a domain and matches a string against the settings configured.
	
	.PARAMETER String
		The string you are wanting to search for.
	
	.PARAMETER ExportPath
		The path you wish to export the effected GPO names in CSV format

	.EXAMPLE
		PS C:\> Search-GPOs -String 'Ask to update automatic links' -ExportPath 'C:\Temp\EffectedGPOS.csv'
	
#>
function Search-GPObjects
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$String,
		[Parameter(Mandatory = $true)]
		[string]$ExportPath
	)
	

	#Create a table to collect the GPO names
	$table = New-Object System.Data.DataTable
	
	[void]$table.Columns.Add("Name")
    [void]$table.Columns.Add("Effected")
	
    # Set the domain to search for GPOs
	$DomainName = $env:USERDNSDOMAIN	
	
    # Find all GPOs in the current domain
	write-host "Finding all the GPOs in $DomainName"
	Import-Module grouppolicy
	$allGposInDomain = Get-GPO -All -Domain $DomainName
	
	# Look through each GPO's XML for the string
	$count=1
	Write-Host "Starting search...."
	foreach ($gpo in $allgposindomain)
	{
		Write-Progress -id 1 -Activity "Processing GPOs" -CurrentOperation $($gpo.DisplayName) -PercentComplete $($count/$($AllGPOsinDomain.count) * 100) -Status "$count of $($AllGPOsinDomain.count)"; $count++
		$report = Get-GPOReport -Guid $gpo.Id -ReportType Xml
		if ($report -match $string)
		{
			write-host "Match found in: $($gpo.DisplayName)"
			[void]$table.Rows.Add($gpo.DisplayName, "Yes")
			
		} # end if
		else
		{
			write-host "No match found in: $($gpo.DisplayName)"
		}
	}
	
    
    $table | Sort-Object -property Name | Export-Csv -Path $ExportPath -NoTypeInformation
}
