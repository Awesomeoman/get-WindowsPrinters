$colComputer = "Server001"

$objOut = @()

foreach ($objComputer in $colComputer) {
    
    $colPrinter = Get-WmiObject win32_printer -ComputerName $objComputer

    foreach ($objPrinter in $colPrinter) {
	
    if ($objPrinter.Name -notlike "Microsoft XPS Document Writer") {
	
    $objPing = "Ping Succeeded"
	
        try {

            $objTest = Test-Connection -ComputerName $objPrinter.PortName -count 1 -ErrorAction Stop

        } catch {

            $objPing = "Ping Failed"
        }
        
            $objOut += [PSCustomObject] @{
          
              "Server" = $objPrinter.SystemName
              "Name" = $objPrinter.Name
              "Location" = $objPrinter.Location
              "Share Name" = $objPrinter.Sharename
              "IP Address" = $objPrinter.PortName
              "Driver" = $objPrinter.DriverName
              "Ping Result" = $objPing            

            }               
        }
    }
}

$objOut | export-csv "D:\Reports\Printer.csv" -NoTypeInformation
