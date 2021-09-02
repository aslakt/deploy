Function Write-Registry {
    <#
    .SYNOPSIS
    Writes to registry
    .DESCRIPTION
    Writes to registry in one single CmdLet
    .PARAMETER Path
    Path to Registry SubKey
    .PARAMETER Name
    Name of Leaf SubKey
    .PARAMETER Value
    Value of Name
    .PARAMETER Type
    Specifies the type of property that this cmdlet adds. The acceptable values for this parameter are:
    - String: Specifies a null-terminated string. Equivalent to REG_SZ.
    - ExpandString: Specifies a null-terminated string that contains unexpanded references to environment variables that are expanded when the value is retrieved. Equivalent to REG_EXPAND_SZ.
    - Binary: Specifies binary data in any form. Equivalent to REG_BINARY.
    - DWord: Specifies a 32-bit binary number. Equivalent to REG_DWORD.
    - MultiString: Specifies an array of null-terminated strings terminated by two null characters. Equivalent to REG_MULTI_SZ.
    - Qword: Specifies a 64-bit binary number. Equivalent to REG_QWORD.
    - Unknown: Indicates an unsupported registry data type, such as REG_RESOURCE_LIST.
    .EXAMPLE
    Write-Registry -Path "HKLM:\SOFTWARE\CustomKey"
    Writes "HKLM:\SOFTWARE\CustomKey" if it does not already excist
    .EXAMPLE
    Write-Registry -Path "HKLM:\SOFTWARE\CustomKey" -Name "CustomKeyName" -Value "SomeValue"
    Writes "HKLM:\SOFTWARE\CustomKey" if it does not already excist and then writes the CustomKeyName Leaf Key and sets its value to "SomeValue"
    .EXAMPLE
    Write-Registry -Path "HKLM:\SOFTWARE\CustomKey" -Name "CustomKeyName" -Value 1 -Type "REG_DWORD"
    Writes "HKLM:\SOFTWARE\CustomKey" if it does not already excist and then writes the CustomKeyName Leaf Key and sets its value to 1 (REG_DWORD)
    #>
    [CmdLetBinding(DefaultParameterSetName="Key")]
    param(
        [Parameter(Mandatory=$true, Position=0, ParameterSetName="Key")]
        [Parameter(Mandatory=$true, Position=0, ParameterSetName="Value")]
        [String]$Path,
        [Parameter(Mandatory=$true, Position=1, ParameterSetName="Value")]
        [String]$Name,
        [Parameter(Mandatory=$true, Position=2, ParameterSetName="Value")]
        [Object]$Value,
        [Parameter(Mandatory=$false, Position=3, ParameterSetName="Value")]
        [ValidateSet("REG_SZ", "REG_EXPAND_SZ", "REG_BINARY", "REG_DWORD", "REG_MULTI_SZ", "REG_QWORD", "REG_RESOURCE_LIST")]
        [string]$Type="REG_SZ",
        [Parameter]
        [Switch]$DefaultUser
    )

    if ($DefaultUser) {
        try {
            REG LOAD HKU\default C:\Users\Default\ntuser.dat
            $Path = "HKU:\default$(Split-Path $Path -NoQualifier)"
        }
        catch {
            Throw "Unable to load default user registry (C:\Users\Default\ntuser.dat)"
        }
    }
    Switch ($PSCmdlet.ParameterSetName) {
        "Key" {
            if (!(Test-Path $Path)) {
                New-Item -Path $Path -Force
            }
        }
        "Value" {
            if (!(Test-Path $Path)) {
                New-Item -Path $Path -Force
            }
            New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force
        }
    }
    if ($DefaultUser) {
        try {
            REG UNLOAD HKU\default
        }
        catch {
            Throw "Unable to unload default user registry (HKU:\default)"
        }
    }
}
