#Author: Jimmy Payne
#Basic script to read the contents of a .xlsx file and see what data is there.
#Reference: https://devblogs.microsoft.com/scripting/grabbing-excel-xlsx-values-with-powershell/
#Required Modules: PSExcel

#To install the required module, do `Install-module PSExcel`

Import-Module PSExcel

$WorkingFile = "$PSScriptRoot\xlsx_sample.xlsx"

foreach ($record in (Import-XLSX -Path $WorkingFile -RowStart 1)){
    Write-Host $record
}

#Note: The -RowSart 1 parameter allows us to skip the header row which contains:

#Product	Category	Price	Stock	Comments

#If everything is in order, you should see the following results:

#@{Product=Cat Food Deluxe; Category=Cat; Price=49.99; Stock=1; Comments=Best cat food ever}
#@{Product=Dog Food Deluxe; Category=Dog; Price=38.48; Stock=56; Comments=Best dog food ever}
#@{Product=Blue-Tounge Food Deluxe; Category=Blue-Tounge Lizard; Price=45.99; Stock=23; Comments=Best Blue-Tounge food ever}
