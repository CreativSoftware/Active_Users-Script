
$list_one = Import-Excel -Path .\ActiveRosterOne.xlsx | Select-Object "Employee Name"
$list_two = Import-Excel -Path .\ActiveRosterTwo.xlsx | Select-Object "Last Name", "First Name"

$new_list_one = @()

foreach ($nameEntry in $list_one) {
    $name = $nameEntry."Employee Name"
    $last, $first = $name -split ","
    
    $first = $first.Trim()
    $last = $last.Trim()

    $newName = "$first $last"

    $obj = New-Object psobject -Property @{
        'FullName' = $newName
    }
    $new_list_one += $obj
}

$combinedList = $list_two | ForEach-Object {
    $fullName = "$($_.'First name') $($_.'Last name')"

    New-Object psobject -Property @{
        'FullName' = $fullName
    }
}

$fullList = $new_list_one += $combinedList

$allDetails = @()
$noResults = @()
foreach ($employee in $fullList){
    $nameDetails = Get-ADUser -Filter "Name -eq '$($employee.FullName)'" -Properties * | Select-Object Name, SamAccountName, UserPrincipalName, Department, Enabled | Where-Object { $_.Enabled -eq $true }

    if ($nameDetails) {
        $allDetails += $nameDetails
    } else {
        $noResults += $employee.FullName
    }   
}

$noResultsfirstname = @()
foreach ($f_name in $noResults){
    $f_name = $f_name.Split(' ')[0]
   
    $firstName = Get-ADUser -Filter "Name -like '*$($f_name)*'" -Properties * | Select-Object Name, SamAccountName 

    $noResultsfirstname += $firstName   
}
    
$username_noresults = @()
foreach($username in $noResultsfirstname){
    $new_noresults = Get-ADUser -Identity $username.SamAccountName -Properties * | Select-Object Name, SamAccountName, UserPrincipalName, Department, Enabled | Where-Object { $_.Enabled -eq $true }
    $username_noresults += $new_noresults
}

$allDetails | Export-Excel -Path .\Results\Results1.xlsx -WorksheetName "Employees"
$username_noresults | Export-Excel -Path .\Results\Results2.xlsx -WorksheetName "Employees"


 

