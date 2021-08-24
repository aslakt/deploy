Function New-WIM2VHD {
    <#
        .SYNOPSIS
        Create a VHD and deploy WIM image to it
        .DESCRIPTION
        Creates a virtual hard disk and deploys the selected WIM image to it.
        Simply attach the VHD to a Hyper-V Virtual Machine and start the VM
        .PARAMETER VhdPath
        Specify where to store the VHD
        .PARAMETER WIMPath
        Specify the path to the WIM image you wish to deploy
        .PARAMETER UnattendPath
        Specify the path to the Unattend.xml file you wish to apply
        .EXAMPLE
        New-SoMVHD .\New-WIM2VHD.ps1 -VhdPath C:\WIM2VHD.vhd -WIMPath "C:\Downloads\image.wim" -UnattendPath "C:\Unattend.xml" -Verbose
        Creates a new VHD deploys the WIM image to it and applies the Unattend definition to it
        .NOTES
        Only support for BIOS for now...
        
        Running the function with the -Verbose switch will provide additional step by step progress information

        ### Authors

        * **aslakt @ GitHub**
    #>
    param(
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateScript({ !(Test-Path $_) })]
        [string]$VhdPath,
        [Parameter(Mandatory=$true,Position=2)]
        [ValidateScript({ Test-Path $_ })]
        [string]$WIMPath,
        [Parameter(Mandatory=$true,Position=2)]
        [ValidateScript({ Test-Path $_ })]
        [string]$UnattendPath
    )

    Write-Host "Creating VHD: $VhdPath" -ForegroundColor Green

    if ($VerbosePreference) { Write-Host " 1/12 - Create, mount and initialize VHD: " -NoNewline -ForegroundColor Yellow }
        $Disk = New-VHD -Path $VhdPath -SizeBytes 32GB -Fixed `
              | Mount-VHD -Passthru `
              | Initialize-Disk -PartitionStyle MBR -Passthru
    if ($VerbosePreference) { Write-Host "DONE" -ForegroundColor Yellow }

    if ($VerbosePreference) { Write-Host " 2/12 - Create System partition: " -NoNewline -ForegroundColor Yellow }
        $BootPartition = (
            $Disk | New-Partition -AssignDriveLetter -Size 499MB -MbrType IFS -IsActive `
                  | Format-Volume -FileSystem NTFS -NewFileSystemLabel "System Reserved" -Confirm:$false -Force
        )
    if ($VerbosePreference) { Write-Host "DONE" -ForegroundColor Yellow }

    if ($VerbosePreference) { Write-Host " 3/12 - Create Windows partition: " -NoNewline -ForegroundColor Yellow }
        $Partition = (
             $Disk | New-Partition -AssignDriveLetter -UseMaximumSize `
                   | Format-Volume -FileSystem NTFS -NewFileSystemLabel "Windows" -Confirm:$false -Force
        )
    if ($VerbosePreference) { Write-Host "DONE" -ForegroundColor Yellow }

  Write-Host "Applying WIM: $WIMPath" -ForegroundColor Green

    if ($VerbosePreference) { Write-Host " 4/12 - Apply SoM image: " -NoNewline -ForegroundColor Yellow }
        Expand-WindowsImage -ImagePath $WIMPath -ApplyPath "$($Partition.DriveLetter):" -Index 1 | Out-Null
    if ($VerbosePreference) { Write-Host "DONE" -ForegroundColor Yellow }

    if ($VerbosePreference) { Write-Host " 5/12 - Install a boot sector: " -NoNewline -ForegroundColor Yellow }
        bootsect /nt60 "$($BootPartition.DriveLetter):" | Out-Null
    if ($VerbosePreference) { Write-Host "DONE" -ForegroundColor Yellow }

    if ($VerbosePreference) { Write-Host " 6/12 - Update boot sector for MBR: " -NoNewline -ForegroundColor Yellow }
        bootsect /nt60 "$($BootPartition.DriveLetter):" /MBR | Out-Null
    if ($VerbosePreference) { Write-Host "DONE" -ForegroundColor Yellow }

    if ($VerbosePreference) { Write-Host " 7/12 - Copy boot loader and create boot record: " -NoNewline -ForegroundColor Yellow }
        bcdboot "$($Partition.DriveLetter):\Windows" /l en-US /s "$($BootPartition.DriveLetter):" /f BIOS | Out-Null
    if ($VerbosePreference) { Write-Host "DONE" -ForegroundColor Yellow }

    if ($VerbosePreference) { Write-Host " 8/12 - Update boot record: " -NoNewline -ForegroundColor Yellow }
        bcdedit /store "$($BootPartition.DriveLetter):\boot\bcd" /timeout 0 | Out-Null
        bcdedit /store "$($BootPartition.DriveLetter):\boot\bcd" /set '{bootmgr}' device locate | Out-Null
        bcdedit /store "$($BootPartition.DriveLetter):\boot\bcd" /set '{default}' device locate | Out-Null
        bcdedit /store "$($BootPartition.DriveLetter):\boot\bcd" /set '{default}' osdevice locate | Out-Null
    if ($VerbosePreference) { Write-Host "DONE" -ForegroundColor Yellow }

    if ($VerbosePreference) { Write-Host " 9/12 - Copy Unattend.xml to $($Partition.DriveLetter):\Windows\Panther: " -NoNewline -ForegroundColor Yellow }
        if (!(Test-Path "$($Partition.DriveLetter):\Windows\Panther")) {
            New-Item "$($Partition.DriveLetter):\Windows\Panther" -ItemType Directory -Force | Out-Null
        }
        Copy-Item -Path $UnattendPath -Destination "$($Partition.DriveLetter):\Windows\Panther\Unattend.xml" -Force
    if ($VerbosePreference) { Write-Host "DONE" -ForegroundColor Yellow }

    if ($VerbosePreference) { Write-Host "10/12 - Apply Unattend.xml: " -NoNewline -ForegroundColor Yellow }
        $tempVerbosePreference = $VerbosePreference
        $VerbosePreference = $false
        Apply-WindowsUnattend -Path "$($Partition.DriveLetter):" -UnattendPath "$($Partition.DriveLetter):\Windows\Panther\Unattend.xml" -NoRestart | Out-Null
        $VerbosePreference = $tempVerbosePreference
    if ($VerbosePreference) { Write-Host "DONE" -ForegroundColor Yellow }

    if ($VerbosePreference) { Write-Host "11/12 - Remove DriveLetter for System partition: " -NoNewline -ForegroundColor Yellow }
        Remove-PartitionAccessPath -DiskNumber $Disk.DiskNumber -PartitionNumber (Get-Partition -DriveLetter $BootPartition.DriveLetter).PartitionNumber -AccessPath "$($BootPartition.DriveLetter):" | Out-Null
    if ($VerbosePreference) { Write-Host "DONE" -ForegroundColor Yellow }

    if ($VerbosePreference) { Write-Host "12/12 - Dismount VHD: " -NoNewline -ForegroundColor Yellow }
        Dismount-VHD -DiskNumber $Disk.DiskNumber | Out-Null
    if ($VerbosePreference) { Write-Host "DONE" -ForegroundColor Yellow }

    if ($VerbosePreference) { Write-Host "Process complete" -ForegroundColor Green }
}
