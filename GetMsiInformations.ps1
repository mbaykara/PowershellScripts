#author Mehmet Ali Baykara
#date : 2019-06-09
#This script is fetching msi information such Product {Name, Code, Version, Upgrade Code} and Manufacturer 

$Dir =  Get-ChildItem -Path C:\MSIs\*.msi -Recurse

foreach ($item in $Dir) { 

    #$Dir = Get-ChildItem -filter { -eq $item} -Property ProductName,ProductCode,ProductVersion,UpgradeCode,Manufacturer

    #$Dir | Get-ChildItem -Object -Property ProductName,ProductCode,ProductVersion,UpgradeCode,Manufacturer

#======================================
#=======    Product Name       ========
#======================================   
try {
    $windowsInstaller = New-Object -com WindowsInstaller.Installer

    $database = $windowsInstaller.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $windowsInstaller, @((Get-Item -Path $item).FullName, 0))

    $view = $database.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $database, ("SELECT Value FROM Property WHERE Property = 'ProductName'"))
    $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)

    $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
    
   echo "Product Name"
    Write-Output -InputObject $($record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1))

} catch {
    Write-Error -Message "Product Name" $_.ToString()s
    
    break
}
 
#======================================
#=======    Product Code       ========
#======================================   
       try {
            $windowsInstaller = New-Object -com WindowsInstaller.Installer

            $database = $windowsInstaller.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $windowsInstaller, @((Get-Item -Path $item).FullName, 0))

            $view = $database.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $database, ("SELECT Value FROM Property WHERE Property = 'ProductCode'"))
            $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)

            $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
            
           echo "Product Code"
            Write-Output -InputObject $($record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1))

        } catch {
            Write-Error -Message "Product Code" $_.ToString()
            
            break
        }

#======================================
#=======    Product Version     ========
#======================================     

try {
    $windowsInstaller = New-Object -com WindowsInstaller.Installer

    $database = $windowsInstaller.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $windowsInstaller, @((Get-Item -Path $item).FullName, 0))

    $view = $database.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $database, ("SELECT Value FROM Property WHERE Property = 'ProductVersion'"))
    $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)

    $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
    
    echo "Product Version"
    Write-Output -InputObject $($record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1))

    $view.GetType().InvokeMember('Close', 'InvokeMethod', $null, $view, $null)
    [Void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($windowsInstaller)
} catch {
     
    Write-Error -Message "Product Version" $_.ToString()
    
    break
}
   
#======================================
#=======    Upgrade Code       ========
#======================================   

try {
    $windowsInstaller = New-Object -com WindowsInstaller.Installer

    $database = $windowsInstaller.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $windowsInstaller, @((Get-Item -Path $item).FullName, 0))

    $view = $database.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $database, ("SELECT Value FROM Property WHERE Property = 'UpgradeCode'"))
    $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)

    $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
    
    echo "Upgrade Code"
    Write-Output -InputObject $($record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1))

    $view.GetType().InvokeMember('Close', 'InvokeMethod', $null, $view, $null)
    [Void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($windowsInstaller)
} catch {
     
    Write-Error -Message "Upgrade Code" $_.ToString()
    
    break
}

#======================================
#=======  Product  Manufacturer =======
#======================================     

try {
    $windowsInstaller = New-Object -com WindowsInstaller.Installer

    $database = $windowsInstaller.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $windowsInstaller, @((Get-Item -Path $item).FullName, 0))

    $view = $database.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $database, ("SELECT Value FROM Property WHERE Property = 'Manufacturer'"))
    $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)

    $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
    
    echo "Manufacturer "
    Write-Output -InputObject $($record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1))

    $view.GetType().InvokeMember('Close', 'InvokeMethod', $null, $view, $null)
    [Void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($windowsInstaller)
} catch {
     
    Write-Error -Message "Manufacturer" $_.ToString()
    
    break
}
}


