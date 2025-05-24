# PowerShell script to parse email addresses in CSV Email column
# and split them into first and last names

# Set the input and output file paths
$inputFile = "Input.csv"
$outputFile = "Input_Updated.csv"

# Import the CSV file
$csvData = Import-Csv $inputFile

# Process each row
foreach ($row in $csvData) {
    # Check if the Email column contains an email address and is in first.last format
    if ($row.Email -match "@" -and $row.Email -match "\.") {
        # Extract the part before the @ symbol
        $emailPrefix = ($row.Email -split "@")[0]
        
        # Split on the dot to get first and last name
        $nameParts = $emailPrefix -split "\."
        
        # If we have at least 2 parts, assign them to first and last
        if ($nameParts.Count -ge 2) {
            $firstName = $nameParts[0]
            $lastName = $nameParts[1]
            
            # Capitalize first letter of each name
            $firstName = (Get-Culture).TextInfo.ToTitleCase($firstName.ToLower())
            $lastName = (Get-Culture).TextInfo.ToTitleCase($lastName.ToLower())
            
            # Create the full name
            $fullName = "$firstName $lastName"
            
            # Add new properties to the row object
            $row | Add-Member -MemberType NoteProperty -Name "First" -Value $firstName -Force
            $row | Add-Member -MemberType NoteProperty -Name "Last" -Value $lastName -Force
            $row | Add-Member -MemberType NoteProperty -Name "FullName_Parsed" -Value $fullName -Force
            
            Write-Host "Processed: $($row.Email) -> First: $firstName, Last: $lastName, FullName: $fullName"
        }
        else {
            # If we can't split properly, add empty fields
            $row | Add-Member -MemberType NoteProperty -Name "First" -Value "" -Force
            $row | Add-Member -MemberType NoteProperty -Name "Last" -Value "" -Force
            $row | Add-Member -MemberType NoteProperty -Name "FullName_Parsed" -Value "" -Force
            Write-Warning "Could not parse name from email: $($row.Email)"
        }
    }
    else {
        # For non-email entries or emails without dots, add empty First and Last fields
        $row | Add-Member -MemberType NoteProperty -Name "First" -Value "" -Force
        $row | Add-Member -MemberType NoteProperty -Name "Last" -Value "" -Force
        $row | Add-Member -MemberType NoteProperty -Name "FullName_Parsed" -Value "" -Force
    }
}

# Export the updated data to a new CSV file
$csvData | Export-Csv $outputFile -NoTypeInformation

Write-Host "`nProcessing complete! Updated file saved as: $outputFile"
Write-Host "New columns 'First', 'Last', and 'FullName_Parsed' have been added to the CSV."