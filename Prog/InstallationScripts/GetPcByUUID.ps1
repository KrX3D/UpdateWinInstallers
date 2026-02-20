
# Get PC by UUID

# Define the mappings between machine GUIDs and names
$UUIDMappings = @{
    "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" = @{
        Name = "EliteBook-G3"
    }
    "YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY" = @{
        Name = "AMD-PC"
        Path = "$path\AMD_PC\*.inf"
    }
}

# Get the machine GUID
#$machineGuid = (Get-WmiObject -Class Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID).ToUpper()
$machineGuid = (Get-CimInstance -Class Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID).ToUpper()

# Check if the machine GUID exists in the mappings
if ($UUIDMappings.ContainsKey($machineGuid)) {
    $uuidMapping = $UUIDMappings[$machineGuid]
    $PCName = $uuidMapping.Name
    
    Write-Verbose ""
    Write-Verbose "PC ist:   $PCName" #-ForegroundColor "Cyan"
    Write-Verbose "GUID ist: $machineGuid" #-ForegroundColor "Cyan"
    Write-Verbose ""

    return $PCName  # Return the driver name as the result
}
else {
    Write-Verbose "UUID does not match" #-ForegroundColor "Red"
    return "Unknown"  # Return "Unknown" if the UUID does not match
}